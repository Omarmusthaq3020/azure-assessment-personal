"""
Azure landing-zone draw.io builder
==================================
Builds the "complete architecture" draw.io diagram for an Azure assessment
workbook (the azure_resources_detailed_*.xlsx export) directly from its sheets:

    frame -> region -> subscription -> VNet -> subnet -> resources

Everything is derived from the data - nothing is specific to one customer:

* Subscriptions are named from the workbook ('Subscription  Name' columns),
  regions from VNet locations, and each region's hub is the VNet with the
  most peerings to other inventoried VNets.
* Resources are drawn inside the subnet they sit in (VMs, secondary NICs,
  private endpoints, AKS + node pools, load balancers, App Gateways,
  firewalls, delegated PostgreSQL/MySQL, App Service VNet integration,
  Container Apps environments, DNS resolver endpoints, Bastion, gateways).
  Anything without a subnet goes into a per-subscription resource strip
  with counts (hover for names).
* Peerings are shown as badges on every VNet, plus routed arrows:
  hub -> each spoke subscription, hub -> other regions, secondary hub ->
  its spokes, external (not inventoried) VNets -> hub, and Internet ->
  Traffic Manager -> the load balancer / App Gateway its endpoint IP is on.
  Arrows only run through empty gutters and lanes; a geometry check
  verifies that no two arrows overlap or cross.

Entry point: build_landing_zone_drawio(file_path) -> draw.io XML string.
"""

import ast
import collections
import ipaddress
import os
import re
import xml.etree.ElementTree as ET

from openpyxl import load_workbook

# ---------------------------------------------------------------- workbook access

SUBNET_RE = re.compile(r"virtualNetworks/([^/'\"\s]+)/subnets/([^'\"/\s,;}\]]+)", re.I)
SUB_ID_RE = re.compile(r"/subscriptions/([0-9a-f-]{36})/", re.I)


def _norm(k):
    return re.sub(r'[^a-z0-9]', '', str(k or '').lower())


class _Row(dict):
    """Sheet row with tolerant column lookup (case / spacing insensitive)."""

    def __init__(self, header, values):
        super().__init__(zip(header, values))
        self._n = {_norm(k): v for k, v in self.items()}
        self.text = ' '.join(str(v) for v in values if v is not None)

    def g(self, *names, default=None):
        for n in names:
            v = self._n.get(_norm(n))
            if v not in (None, '', 'None'):
                return v
        return default


class _Book:
    def __init__(self, path):
        self.wb = load_workbook(path, read_only=True, data_only=True)
        self.names = self.wb.sheetnames
        self._cache = {}

    def sheet(self, name):
        if name not in self._cache and name not in self.names:
            # newer exports split types that share a friendly name: 'Connections' -> 'Connections (network)', ...
            variants = [n for n in self.names if n.startswith(f'{name} (')]
            if variants:
                self._cache[name] = [r for v in variants for r in self.sheet(v)]
        if name not in self._cache:
            rows = []
            if name in self.names:
                it = self.wb[name].iter_rows(values_only=True)
                header = next(it, None)
                if header:
                    rows = [_Row(header, r) for r in it if r and r[0] not in (None, '')]
            self._cache[name] = rows
        return self._cache[name]


def _subnets_of(row):
    seen = []
    for m in SUBNET_RE.findall(row.text):
        k = (m[0].lower(), m[1].lower())
        if k not in seen:
            seen.append(k)
    return seen


def _as_list(v):
    if v in (None, '', 'None'):
        return []
    if isinstance(v, (list, tuple)):
        return list(v)
    try:
        x = ast.literal_eval(str(v))
        return list(x) if isinstance(x, (list, tuple)) else [str(x)]
    except (ValueError, SyntaxError):
        return [s.strip() for s in re.split(r'[;,]', str(v)) if s.strip()]


REGION_NAMES = {
    'centralindia': 'Central India', 'southindia': 'South India', 'westindia': 'West India',
    'eastus': 'East US', 'eastus2': 'East US 2', 'westus': 'West US', 'westus2': 'West US 2',
    'westeurope': 'West Europe', 'northeurope': 'North Europe', 'uksouth': 'UK South',
    'southeastasia': 'Southeast Asia', 'eastasia': 'East Asia', 'uaenorth': 'UAE North',
}


def _region_title(loc):
    return REGION_NAMES.get(loc, loc)


def _strip_common_prefix(names):
    """['vm-cns-vmprod-app2', 'vm-cns-vmprod-db1'] -> ['app2', 'db1'] (only when 2+ names share a prefix)."""
    if len(names) < 2:
        return names
    pre = os.path.commonprefix(names)
    pre = pre[:max(pre.rfind('-'), pre.rfind('_')) + 1]      # cut at a separator, never mid-word
    return [n[len(pre):] or n for n in names] if pre else names


def _lb_summary(r, pip_ip, nic_vm):
    """Frontend, rules and backends of a load balancer row: a short 2-line label and a full tooltip."""
    try:
        fe = ast.literal_eval(str(r.g('Frontend IP Configurations') or '[]'))
        rules = ast.literal_eval(str(r.g('Load Balancing Rules') or '[]'))
        pools = ast.literal_eval(str(r.g('Backend Address Pools') or '[]'))
        probes = {p['name']: p.get('properties', {}) for p in ast.literal_eval(str(r.g('Probes') or '[]'))}
    except (ValueError, SyntaxError):
        return {'short': '', 'tip': None}
    ips = []
    for f in fe:
        p = f.get('properties', {})
        if p.get('privateIPAddress'):
            ips.append(f"internal {p['privateIPAddress']}")
        pip = (p.get('publicIPAddress') or {}).get('id', '').split('/')[-1]
        if pip:
            ips.append(f"public {pip_ip.get(pip) or pip}")
    backends, orphans = [], []          # orphans: backend NICs whose VM is not in the inventory
    for pool in pools:
        for c in pool.get('properties', {}).get('backendIPConfigurations', []):
            nic = c.get('id', '').split('/networkInterfaces/')[-1].split('/')[0].lower()
            if nic in nic_vm:
                if nic_vm[nic] not in backends:
                    backends.append(nic_vm[nic])
            elif nic and nic not in orphans:
                orphans.append(nic)
    lines, ports = [], []
    for rule in rules:
        p = rule.get('properties', {})
        probe = probes.get((p.get('probe') or {}).get('id', '').split('/')[-1], {})
        ports.append(str(p.get('frontendPort')))
        lines.append(f"{rule.get('name')}: {p.get('protocol')} {p.get('frontendPort')} -> {p.get('backendPort')}"
                     + (f"  (health: {probe.get('protocol')} :{probe.get('port')}{probe.get('requestPath') or ''})" if probe else ''))
    short_backends = _strip_common_prefix(backends)
    shown = ', '.join(short_backends[:4]) + (f' +{len(backends) - 4}' if len(backends) > 4 else '')
    short = f"({' · '.join(ips)})" if ips else ''
    if ports and backends:
        short += f"<br><font color='#555555'>{', '.join(ports[:4])} → {shown}</font>"
    elif ports and orphans:
        # rules exist but every backend NIC belongs to a VM that no longer exists: traffic goes nowhere
        short += f"<br><font color='#C0392B'>{', '.join(ports[:4])} → no VM ({len(orphans)} orphan NICs)</font>"
    tip = '\n'.join([f"Frontend: {', '.join(ips) or 'n/a'}", 'Rules:'] + [f'  {l}' for l in lines] +
                    [f"Backends: {', '.join(backends) or 'none'}"] +
                    ([f"Backend NICs with no VM: {', '.join(orphans)}"] if orphans else []))
    return {'short': short, 'tip': tip}


# ---------------------------------------------------------------- icons (all exist in draw.io's azure2 library)

def _icon(p):
    return f'img/lib/azure2/{p}.svg'


PE_ICON = {
    'microsoft.web/sites': 'compute/Function_Apps', 'microsoft.datafactory/factories': 'databases/Data_Factory',
    'microsoft.dbforpostgresql/flexibleservers': 'databases/Azure_Database_PostgreSQL_Server',
    'microsoft.dbformysql/flexibleservers': 'databases/Azure_Database_MySQL_Server',
    'microsoft.cache/redisenterprise': 'databases/Cache_Redis', 'microsoft.cache/redis': 'databases/Cache_Redis',
    'microsoft.storage/storageaccounts': 'storage/Storage_Accounts', 'microsoft.signalrservice/webpubsub': 'web/SignalR',
    'microsoft.signalrservice/signalr': 'web/SignalR', 'microsoft.eventhub/namespaces': 'analytics/Event_Hubs',
    'microsoft.servicebus/namespaces': 'integration/Service_Bus', 'microsoft.keyvault/vaults': 'security/Key_Vaults',
    'microsoft.containerregistry/registries': 'containers/Container_Registries',
    'microsoft.insights/privatelinkscopes': 'management_governance/Monitor', 'microsoft.sql/servers': 'databases/SQL_Server',
    'microsoft.sql/managedinstances': 'databases/SQL_Managed_Instance', 'microsoft.documentdb/databaseaccounts': 'databases/Azure_Cosmos_DB',
    'microsoft.search/searchservices': 'app_services/Search_Services', 'microsoft.cognitiveservices/accounts': 'ai_machine_learning/Cognitive_Services',
    'microsoft.databricks/workspaces': 'analytics/Azure_Databricks', 'microsoft.apimanagement/service': 'integration/API_Management_Services',
    'microsoft.synapse/workspaces': 'analytics/Azure_Synapse_Analytics', 'microsoft.eventgrid/topics': 'integration/Event_Grid_Topics',
}

# resource strip: workbook 'Type' (= sheet name) -> (label, icon); unknown types get a generic icon
STRIP = collections.OrderedDict([
    ('Storage Accounts', ('Storage accounts', 'storage/Storage_Accounts')),
    ('Key Vault', ('Key Vaults', 'security/Key_Vaults')),
    ('Redis Enterprise', ('Azure Managed Redis', 'databases/Cache_Redis')),
    ('Redis Caches', ('Azure Cache for Redis', 'databases/Cache_Redis')),
    ('PostgreSQL Flexible Servers', ('PostgreSQL flexible servers', 'databases/Azure_Database_PostgreSQL_Server')),
    ('MySQL Flexible Servers', ('MySQL flexible servers', 'databases/Azure_Database_MySQL_Server')),
    ('Servers', ('Azure SQL servers', 'databases/SQL_Server')),
    ('Databases', ('SQL databases', 'databases/SQL_Database')),
    ('Cosmos DB', ('Cosmos DB', 'databases/Azure_Cosmos_DB')),
    ('Namespaces', ('Event Hubs / Service Bus namespaces', 'analytics/Event_Hubs')),
    ('Azure Container Registry', ('Container registries', 'containers/Container_Registries')),
    ('AKS', ('AKS clusters', 'containers/Kubernetes_Services')),
    ('Virtual Machines', ('Virtual machines', 'compute/Virtual_Machine')),
    ('App Service', ('App Services / Functions', 'containers/App_Services')),
    ('App Service Plan', ('App Service plans', 'containers/App_Services')),
    ('Workflows', ('Logic App workflows', 'integration/Logic_Apps')),
    ('Data Factories', ('Data Factory', 'databases/Data_Factory')),
    ('Web Pub Sub', ('Web PubSub', 'web/SignalR')),
    ('Container Apps', ('Container Apps', 'other/Worker_Container_App')),
    ('Container Groups', ('Container instances', 'containers/Container_Instances')),
    ('Cognitive Services', ('AI services / Foundry', 'ai_machine_learning/Cognitive_Services')),
    ('Search Services', ('AI Search', 'app_services/Search_Services')),
    ('Capacities', ('Fabric capacity', 'analytics/Power_BI_Embedded')),
    ('Grafana', ('Managed Grafana', 'other/Grafana')),
    ('Automation Accounts', ('Automation accounts', 'management_governance/Automation_Accounts')),
    ('Bastion Hosts', ('Bastion', 'networking/Bastions')),
    ('Application Gateways', ('Application Gateways', 'networking/Application_Gateways')),
    ('Load Balancers', ('Load balancers', 'networking/Load_Balancers')),
    ('Trafficmanagerprofiles', ('Traffic Manager', 'networking/Traffic_Manager_Profiles')),
    ('Workspaces', ('Log Analytics workspaces', 'analytics/Log_Analytics_Workspaces')),
    ('Public IP Addresses', ('Public IPs', 'networking/Public_IP_Addresses')),
    ('Private Dns Zones', ('Private DNS zones', 'networking/DNS_Zones')),
    ('User Assigned Identities', ('Managed identities', 'identity/Managed_Identities')),
    ('Network Security Groups', ('NSGs', 'networking/Network_Security_Groups')),
    ('Route Tables', ('Route tables (UDR)', 'networking/Route_Tables')),
    ('Loadtests', ('Azure Load Testing', 'other/Azure_Load_Testing')),
    ('Recovery Services Vault', ('Recovery Services vaults', 'management_governance/Recovery_Services_Vaults')),
    ('Profiles', ('Front Door / CDN profiles', 'networking/Front_Doors')),
    ('Afdendpoints', ('Front Door endpoints', 'networking/Front_Doors')),
    ('Frontdoor Web Application Firewall Policies', ('Front Door WAF policies', 'networking/Web_Application_Firewall_Policies_WAF')),
    ('Application Gateway Web Application Firewall Policies', ('App Gateway WAF policies', 'networking/Web_Application_Firewall_Policies_WAF')),
    ('Firewall Policies', ('Firewall policies', 'networking/Azure_Firewall_Policy')),
    ('Disk Encryption Sets', ('Disk encryption sets', 'compute/Disk_Encryption_Sets')),
])
# drawn elsewhere (inside subnets or as VNets) or too granular to be useful in the strip
STRIP_SKIP = {'Virtual Networks', 'Network Interfaces', 'Private Endpoints', 'Virtual Machine Scale Sets', 'Disks',
              'Snapshots', 'Virtual Network Links', 'Extensions', 'Network Watchers', 'Prometheus Rule Groups',
              'Smart Detector Alert Rules', 'Metric Alerts', 'Action Group', 'Data Collection Rules',
              'Data Collection Endpoints', 'Runbooks', 'Schedules', 'Connections', 'Dashboards', 'Inbound Endpoints',
              'Outbound Endpoints', 'Dns Forwarding Rulesets', 'Certificates', 'Accounts', 'Account', 'Builders',
              'Integration Fabrics', 'Jobs', 'Projects', 'Managed Environments', 'Private Link Scopes', 'Dns Resolvers',
              'Firewalls', 'Activity Log Alerts', 'Restore Point Collections', 'Images', 'Versions', 'Galleries',
              'Ssh Public Keys', 'Disk Accesses', 'Autoscalesettings', 'Discovery Sources', 'Assessment Projects',
              'Master Sites', 'Server Sites', 'V Mware Sites', 'My S Q L Sites', 'P G S Q L Sites', 'Storagesites',
              'Sites', 'Vmware Sites', 'Querypacks', 'Flow Logs', 'Scheduled Query Rules', 'Action Rules',
              'Maintenance Configurations', 'Run Commands'}

# resource strip by ARM type (sheet names are ambiguous: 'Connections' is both VPN/ER and Logic App API connections)
ARM_STRIP = {
    'microsoft.network/expressroutecircuits': ('ExpressRoute circuits', 'networking/ExpressRoute_Circuits'),
    'microsoft.network/virtualnetworkgateways': ('VPN / ExpressRoute gateways', 'networking/Virtual_Network_Gateways'),
    'microsoft.network/ddosprotectionplans': ('DDoS protection plans', 'networking/DDoS_Protection_Plans'),
    'microsoft.network/privatelinkservices': ('Private Link services', 'networking/Private_Link_Service'),
    'microsoft.network/firewallpolicies': ('Firewall policies', 'networking/Azure_Firewall_Policy'),
    'microsoft.network/applicationgatewaywebapplicationfirewallpolicies': ('App Gateway WAF policies', 'networking/Web_Application_Firewall_Policies_WAF'),
    'microsoft.network/frontdoorwebapplicationfirewallpolicies': ('Front Door WAF policies', 'networking/Web_Application_Firewall_Policies_WAF'),
    'microsoft.network/natgateways': ('NAT gateways', 'networking/NAT'),
    'microsoft.network/bastionhosts': ('Bastion', 'networking/Bastions'),
    'microsoft.network/dnszones': ('Public DNS zones', 'networking/DNS_Zones'),
    'microsoft.network/privatednszones': ('Private DNS zones', 'networking/DNS_Zones'),
    'microsoft.network/trafficmanagerprofiles': ('Traffic Manager', 'networking/Traffic_Manager_Profiles'),
    'microsoft.network/loadbalancers': ('Load balancers', 'networking/Load_Balancers'),
    'microsoft.network/applicationgateways': ('Application Gateways', 'networking/Application_Gateways'),
    'microsoft.compute/availabilitysets': ('Availability sets', 'compute/Availability_Sets'),
    'microsoft.compute/virtualmachines': ('Virtual machines', 'compute/Virtual_Machine'),
    'microsoft.compute/galleries': ('Compute galleries', 'compute/Azure_Compute_Galleries'),
    'microsoft.compute/galleries/applications': ('VM application definitions', 'other/VM_Application_Definition'),
    'microsoft.compute/diskencryptionsets': ('Disk encryption sets (CMK)', 'compute/Disk_Encryption_Sets'),
    'microsoft.sqlvirtualmachine/sqlvirtualmachines': ('SQL Server on VMs', 'databases/Azure_SQL_VM'),
    'microsoft.sql/servers': ('Azure SQL servers', 'databases/SQL_Server'),
    'microsoft.sql/servers/databases': ('SQL databases', 'databases/SQL_Database'),
    'microsoft.sql/managedinstances': ('SQL managed instances', 'databases/SQL_Managed_Instance'),
    'microsoft.documentdb/databaseaccounts': ('Cosmos DB accounts', 'databases/Azure_Cosmos_DB'),
    'microsoft.cache/redis': ('Azure Cache for Redis', 'databases/Cache_Redis'),
    'microsoft.cache/redisenterprise': ('Azure Managed Redis', 'databases/Cache_Redis'),
    'microsoft.dbforpostgresql/flexibleservers': ('PostgreSQL flexible servers', 'databases/Azure_Database_PostgreSQL_Server'),
    'microsoft.dbformysql/flexibleservers': ('MySQL flexible servers', 'databases/Azure_Database_MySQL_Server'),
    'microsoft.web/sites': ('App Services / Functions', 'containers/App_Services'),
    'microsoft.web/sites/slots': ('App Service deployment slots', 'general/Web_Slots'),
    'microsoft.web/serverfarms': ('App Service plans', 'app_services/App_Service_Plans'),
    'microsoft.web/connections': ('Logic App API connections', 'devops/API_Connections'),
    'microsoft.web/connectiongateways': ('On-premises data gateways', 'networking/On_Premises_Data_Gateways'),
    'microsoft.logic/workflows': ('Logic App workflows', 'integration/Logic_Apps'),
    'microsoft.apimanagement/service': ('API Management', 'integration/API_Management_Services'),
    'microsoft.eventgrid/systemtopics': ('Event Grid system topics', 'integration/System_Topic'),
    'microsoft.eventgrid/topics': ('Event Grid topics', 'integration/Event_Grid_Topics'),
    'microsoft.servicebus/namespaces': ('Service Bus namespaces', 'integration/Service_Bus'),
    'microsoft.eventhub/namespaces': ('Event Hubs namespaces', 'analytics/Event_Hubs'),
    'microsoft.databricks/workspaces': ('Databricks workspaces', 'analytics/Azure_Databricks'),
    'microsoft.databricks/accessconnectors': ('Databricks access connectors', 'analytics/Azure_Databricks'),
    'microsoft.machinelearningservices/workspaces': ('ML workspaces', 'ai_machine_learning/Machine_Learning'),
    'microsoft.machinelearningservices/workspaces/onlineendpoints': ('ML online endpoints', 'ai_machine_learning/Machine_Learning'),
    'microsoft.machinelearningservices/workspaces/onlineendpoints/deployments': ('ML endpoint deployments', 'ai_machine_learning/Machine_Learning'),
    'microsoft.cognitiveservices/accounts': ('AI services / Foundry', 'ai_machine_learning/Cognitive_Services'),
    'microsoft.search/searchservices': ('AI Search', 'app_services/Search_Services'),
    'microsoft.insights/workbooks': ('Workbooks', 'general/Workbooks'),
    'microsoft.insights/components': ('Application Insights', 'devops/Application_Insights'),
    'microsoft.operationsmanagement/solutions': ('Log Analytics solutions', 'management_governance/Solutions'),
    'microsoft.resourcegraph/queries': ('Resource Graph queries', 'management_governance/Resource_Graph_Explorer'),
    'microsoft.dataprotection/backupvaults': ('Backup vaults', 'other/Backup_Vault'),
    'microsoft.recoveryservices/vaults': ('Recovery Services vaults', 'management_governance/Recovery_Services_Vaults'),
    'microsoft.automation/automationaccounts': ('Automation accounts', 'management_governance/Automation_Accounts'),
    'microsoft.security/automations': ('Defender workflow automations', 'security/Azure_Defender'),
    'microsoft.saas/resources': ('Marketplace SaaS resources', 'general/Marketplace'),
    'microsoft.containerregistry/registries': ('Container registries', 'containers/Container_Registries'),
    'microsoft.containerservice/managedclusters': ('AKS clusters', 'containers/Kubernetes_Services'),
    'microsoft.storage/storageaccounts': ('Storage accounts', 'storage/Storage_Accounts'),
    'microsoft.keyvault/vaults': ('Key Vaults', 'security/Key_Vaults'),
    'microsoft.datafactory/factories': ('Data Factory', 'databases/Data_Factory'),
    'microsoft.dashboard/grafana': ('Managed Grafana', 'other/Grafana'),
    'microsoft.app/containerapps': ('Container Apps', 'other/Worker_Container_App'),
    'microsoft.app/managedenvironments': ('Container Apps environments', 'other/Container_App_Environments'),
    'microsoft.relay/namespaces': ('Relay namespaces', 'integration/Relays'),
    'microsoft.containerinstance/containergroups': ('Container instances', 'containers/Container_Instances'),
}
# never in the strip: system / child objects with no architectural meaning
ARM_SKIP = {'microsoft.network/networkintentpolicies', 'microsoft.containerregistry/registries/webhooks',
            'microsoft.compute/virtualmachines/extensions', 'microsoft.network/networkwatchers',
            'microsoft.network/networkwatchers/flowlogs'}
# drawn elsewhere: in subnets, as markers, or in the External column panels
ARM_HANDLED = {'microsoft.cdn/profiles', 'microsoft.cdn/profiles/afdendpoints', 'microsoft.cdn/profiles/origingroups/origins',
               'microsoft.cdn/profiles/customdomains', 'microsoft.cdn/profiles/afdendpoints/routes',
               'microsoft.cdn/profiles/endpoints', 'microsoft.network/routetables', 'microsoft.network/connections',
               'microsoft.network/localnetworkgateways', 'microsoft.network/expressroutecircuits',
               'microsoft.insights/datacollectionrules', 'microsoft.insights/datacollectionruleassociations',
               'microsoft.insights/diagnosticsettings', 'microsoft.authorization/policyassignments', 'microsoft.authorization/locks',
               'microsoft.security/pricings', 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems',
               'microsoft.insights/components'}


def _arm_label(arm):
    """Readable fallback for an unmapped ARM type: 'microsoft.foo/barBazs' -> 'Bar bazs (foo)'."""
    ns, _, leaf = arm.partition('/')
    leaf = re.sub(r'(?<!^)(?=[A-Z])', ' ', leaf.split('/')[-1]).replace('_', ' ').strip().capitalize()
    return f"{leaf} ({ns.replace('microsoft.', '')})"


# ---------------------------------------------------------------- styles

BOX = 'rounded=1;arcSize=2;whiteSpace=wrap;html=1;verticalAlign=top;align=left;spacingLeft=6;spacingTop=2;container=1;collapsible=0;'
ICON = ('image;aspect=fixed;html=1;points=[];align=center;fontSize=9;image={};verticalLabelPosition=bottom;'
        'verticalAlign=top;labelBackgroundColor=none;whiteSpace=wrap;labelWidth=96;')
MARKER = 'image;aspect=fixed;html=1;points=[];image={};labelBackgroundColor=none;'
STALE_KINDS = {'networking/Network_Security_Groups': 'NSGs not attached', 'networking/Public_IP_Addresses': 'Public IPs not attached',
               'networking/Network_Interfaces': 'NICs not attached to a VM', 'compute/Disks': 'Unattached disks',
               'identity/Managed_Identities': 'Identities not assigned'}
MARKER_RIGHT_LABEL = ('image;aspect=fixed;html=1;points=[];image={};labelBackgroundColor=none;fontSize=8;fontColor=#555555;'
                      'labelPosition=right;verticalLabelPosition=middle;align=left;verticalAlign=middle;spacingLeft=2;')
MARKER_LEFT_LABEL = ('image;aspect=fixed;html=1;points=[];image={};labelBackgroundColor=none;fontSize=8;fontColor=#555555;'
                     'labelPosition=left;verticalLabelPosition=middle;align=right;verticalAlign=middle;spacingRight=2;')
ICON_RIGHT_LABEL = ICON.replace('verticalLabelPosition=bottom;verticalAlign=top;',
                                'labelPosition=right;verticalLabelPosition=middle;align=left;verticalAlign=middle;')
SUBNET_ON = BOX + 'fillColor=#FFFFFF;strokeColor=#7FA9D9;dashed=1;fontSize=10;'
SUBNET_OFF = BOX + 'fillColor=#FFFFFF;strokeColor=#C8D6E5;dashed=1;fontSize=10;fontColor=#8A8A8A;'
VNET_HUB = 'fillColor=#F3EEFA;strokeColor=#7B4FB5;strokeWidth=2;'
VNET_FW = 'fillColor=#FFF7F7;strokeColor=#C0392B;strokeWidth=1.5;'
VNET_ISOLATED = 'fillColor=#F5F5F5;strokeColor=#9E9E9E;strokeWidth=1;dashed=1;'
VNET_SPOKE = 'fillColor=#EAF3FC;strokeColor=#1F6FC5;strokeWidth=1.5;'
SUB_STYLE = BOX + 'fillColor=#FFFFFF;strokeColor=#1F6FC5;strokeWidth=2;fontStyle=1;fontSize=14;fontColor=#1F6FC5;'
REGION_STYLE = BOX + 'fillColor=#FAFAFA;strokeColor=#5D6D7E;strokeWidth=1.5;fontStyle=1;fontSize=15;'
EXT_STYLE = BOX + 'fillColor=#FFFFFF;strokeColor=#9E9E9E;dashed=1;fontStyle=1;fontSize=12;'
BADGE_COLORS = ['#7B4FB5', '#2E86C1', '#1F6FC5', '#C55A11', '#E67E22', '#16A085', '#8E44AD', '#B03A2E']

CW, RH, PER_ROW = 100, 92, 4          # icon cell width, icon row height, icons per subnet row
SW = 16 + PER_ROW * CW                # subnet width
GAP = 12                              # gap between subnets / paired VNets
COL_GAP = 50                          # gap between VNet columns in a subscription
MAX_ROW_W = 7600                      # wrap subscriptions of the primary region into a second row past this width
TARGET_SUB_H = 1600                   # a subscription box grows extra VNet columns rather than past this height
TARGET_VNET_H = 1400                  # a VNet grows extra subnet columns rather than past this height
LANDSCAPE = 0.6                       # ...and a subscription box aims to be no taller than 0.6 x its width
LANE_STEP = 28                        # spacing of parallel edge lanes


class _Diagram:
    """mxGraph document with absolute-position tracking (used to route edges)."""

    def __init__(self, name):
        self.mxfile = ET.Element('mxfile', host='Electron', version='28.2.8')
        diagram = ET.SubElement(self.mxfile, 'diagram', name=name, id='d1')
        model = ET.SubElement(diagram, 'mxGraphModel', dx='1188', dy='729', grid='1', gridSize='10', guides='1',
                              tooltips='1', connect='1', arrows='1', fold='1', page='0', pageScale='1',
                              pageWidth='850', pageHeight='1100', math='0', shadow='0')
        self.root = ET.SubElement(model, 'root')
        ET.SubElement(self.root, 'mxCell', id='0')
        ET.SubElement(self.root, 'mxCell', id='1', parent='0')
        self.n = 0
        self.abs = {'1': (0, 0, 0, 0)}
        self.routes = []

    def nid(self):
        self.n += 1
        return f'c{self.n}'

    def cell(self, parent, value, style, x, y, w, h, tooltip=None):
        cid = self.nid()
        if tooltip:
            uo = ET.SubElement(self.root, 'UserObject', id=cid, label=value, tooltip=tooltip)
            c = ET.SubElement(uo, 'mxCell', style=style, parent=parent, vertex='1')
        else:
            c = ET.SubElement(self.root, 'mxCell', id=cid, value=value, style=style, parent=parent, vertex='1')
        ET.SubElement(c, 'mxGeometry', x=str(round(x)), y=str(round(y)), width=str(round(w)), height=str(round(h)),
                      **{'as': 'geometry'})
        px, py = self.abs[parent][:2]
        self.abs[cid] = (px + x, py + y, w, h)
        return cid

    def anchor(self, cid, rel):
        x, y, w, h = self.abs[cid]
        rel = rel or (0.5, 0.5)
        return (x + rel[0] * w, y + rel[1] * h)

    def rel_y(self, cid, y):
        _, cy, _, ch = self.abs[cid]
        return (y - cy) / ch

    def edge(self, src, tgt, label='', extra='', exit=None, entry=None, pts=(), label_pos=None):
        """Edge leaving src at relative point exit, through absolute waypoints pts, into tgt at entry.
        Callers keep every segment horizontal or vertical so the edge stays in its gutter.
        label_pos: where the label sits along the edge, -1 (source) .. 1 (target); default the middle."""
        style = ('edgeStyle=none;rounded=1;html=1;endArrow=block;endFill=1;strokeWidth=1.5;'
                 'fontSize=10;labelBackgroundColor=#FFFFFF;' + extra)
        for key, p in (('exit', exit), ('entry', entry)):
            if p:
                style += f'{key}X={p[0]:.4f};{key}Y={p[1]:.4f};{key}Dx=0;{key}Dy=0;{key}Perimeter=0;'
        c = ET.SubElement(self.root, 'mxCell', id=self.nid(), value=label, style=style,
                          parent='1', source=src, target=tgt, edge='1')
        g = ET.SubElement(c, 'mxGeometry', relative='1', **{'as': 'geometry'})
        if label_pos is not None:
            g.set('x', f'{label_pos:.2f}')
        if pts:
            arr = ET.SubElement(g, 'Array', **{'as': 'points'})
            for x, y in pts:
                ET.SubElement(arr, 'mxPoint', x=str(round(x)), y=str(round(y)))
        self.routes.append((re.sub('<[^>]+>', ' ', label), [self.anchor(src, exit), *pts, self.anchor(tgt, entry)]))

    @staticmethod
    def _segs(p):
        return list(zip(p, p[1:]))

    @staticmethod
    def _hit(s1, s2):
        """True if two axis-aligned segments overlap or intersect (shared endpoints excepted)."""
        (a, b), (c, d) = s1, s2
        ax0, ax1 = sorted((a[0], b[0])); ay0, ay1 = sorted((a[1], b[1]))
        cx0, cx1 = sorted((c[0], d[0])); cy0, cy1 = sorted((c[1], d[1]))
        touch = ax0 <= cx1 and cx0 <= ax1 and ay0 <= cy1 and cy0 <= ay1
        return touch and not ({a, b} & {c, d})

    def would_conflict(self, start, pts, end):
        """Would an edge along start -> pts -> end touch an edge already drawn?"""
        new = self._segs([start, *pts, end])
        return any(self._hit(s, t) for _, p in self.routes for s in self._segs(p) for t in new)

    def conflicts(self):
        """Pairs of edges whose axis-aligned segments touch (shared endpoints excepted), and diagonal segments."""
        segs, hit = self._segs, self._hit
        bad = []
        for i, (l1, p1) in enumerate(self.routes):
            for l2, p2 in self.routes[i + 1:]:
                if any(hit(s, t) for s in segs(p1) for t in segs(p2)):
                    bad.append(f'{l1[:40]} x {l2[:40]}')
            for a, b in segs(p1):
                if abs(a[0] - b[0]) > 0.5 and abs(a[1] - b[1]) > 0.5:
                    bad.append(f'{l1[:40]}: diagonal segment')
        return bad

    def xml(self):
        ET.indent(self.mxfile)
        return ET.tostring(self.mxfile, encoding='unicode')


def _badge(text, color):
    return (f'<span style="background-color:{color};color:#FFFFFF;font-size:9px;font-weight:bold;'
            f'padding:0px 3px;border-radius:3px;white-space:nowrap;">{text}</span>')


def _short(name):
    s = re.sub(r'(?i)[-_]?vnet[-_]?\d*$', '', name)
    s = re.sub(r'(?i)^vnet[-_]', '', s)
    return (s or name).upper()[:18]


# ---------------------------------------------------------------- inventory model

class _Inventory:
    def __init__(self, book):
        self.b = book
        self._subscriptions()
        self._vnets()
        self.place = collections.defaultdict(list)   # (vnet_l, subnet_l) -> [(label, icon)]
        self.placed = set()                          # (type, name_l) drawn inside a subnet
        self.item_res = {}                           # (vnet_l, subnet_l, label) -> (type, name_l) of the drawn resource
        self.markers = collections.defaultdict(list) # (type, name_l) -> [(icon, tooltip)] small icons beside a resource
        self.subnet_nsg = {}                         # (vnet_l, subnet_l) -> NSG name
        self.res_vnet = {}                           # ('res', type, name_l) -> vnet name (for edges)
        self.pe_targets = set()                      # names of resources reached through a private endpoint
        self.private_aks = set()
        self.unplaced_pe = []
        self.stale = collections.defaultdict(list)   # subscription id -> [(icon, name, reason)]
        self._disks()
        self._place_resources()
        self._attachments()
        self._routes()
        self._dns_links()
        self._links()
        self._frontdoor()
        self._hybrid()
        self.markers_lock = {}
        self._governance()
        self._strip()

    # subscriptions: id -> display name, from any row that carries both
    def _subscriptions(self):
        self.sub_name = {}
        all_ids = collections.Counter()
        for r in self.b.sheet('AllResources'):
            sid = str(r.g('Subscription') or '').lower()
            if re.fullmatch(r'[0-9a-f-]{36}', sid):
                all_ids[sid] += 1
        for sheet in self.b.names:
            if len(self.sub_name) >= len(all_ids):
                break
            for r in self.b.sheet(sheet)[:200]:
                name = r.g('Subscription Name')
                ids = SUB_ID_RE.findall(r.text)
                if name and ids:
                    self.sub_name.setdefault(ids[0].lower(), str(name))
        for r in self.b.sheet('Subscriptions'):          # every assessed subscription, including empty ones
            sid = str(r.g('Subscription Id') or '').lower()
            if re.fullmatch(r'[0-9a-f-]{36}', sid):
                self.sub_name[sid] = str(r.g('Subscription Name') or self.sub_name.get(sid) or f'Subscription {sid[:8]}')
                all_ids.setdefault(sid, 0)
        for sid in all_ids:
            self.sub_name.setdefault(sid, f'Subscription {sid[:8]}')
        self.name_to_id = {v: k for k, v in self.sub_name.items()}
        self.sub_count = all_ids

    def sub_of_row(self, r):
        name = r.g('Subscription Name')
        if name and str(name) in self.name_to_id:
            return self.name_to_id[str(name)]
        ids = SUB_ID_RE.findall(r.text)
        return ids[0].lower() if ids else None

    def _vnets(self):
        self.vnets = collections.OrderedDict()
        for r in self.b.sheet('Virtual Networks'):
            names = [s.strip() for s in str(r.g('Subnet Names') or '').split(';') if s.strip()]
            prefs = [s.strip() for s in str(r.g('Subnet Prefixes') or '').split(';')]
            # exports carry a summary column ("peer-x (Connected) -> remote; ...") and a raw JSON column whose
            # name differs only in spacing; take remotes from both, never the '.../virtualNetworkPeerings' type
            peers = []
            for k, val in r.items():
                if _norm(k) == 'virtualnetworkpeerings' and val:
                    txt = str(val)
                    peers += re.findall(r'-> ([^;]+?)(?:;|$)', txt)
                    peers += re.findall(r"remoteVirtualNetwork'?\"?:\s*\{'?\"?id'?\"?:\s*'?\"?[^'\"]*/virtualNetworks/([^'\"/]+)", txt)
            self.vnets[r['Name']] = dict(
                name=r['Name'], loc=str(r.g('Location') or 'unknown').lower(), sub=self.sub_of_row(r),
                space=[str(x) for x in _as_list(r.g('Address Prefixes', 'Address Space'))],
                dns=[str(x) for x in _as_list(r.g('Dns Servers'))],
                subnets=names, prefix=dict(zip(names, prefs)),
                peers=sorted(set(p.strip() for p in peers if p.strip() and p.strip() != r['Name'])))
        self.vlow = {n.lower(): n for n in self.vnets}

    def put(self, vn, sn, label, ic, rtype=None, rname=None):
        key = (vn.lower(), sn.lower())
        if (label, _icon(ic)) not in self.place[key]:
            self.place[key].append((label, _icon(ic)))
        if rtype:
            self.placed.add((rtype, str(rname).lower()))
            self.item_res[(key[0], key[1], label)] = (rtype, str(rname).lower())
            self.res_vnet[('res', rtype, str(rname).lower())] = self.vlow.get(vn.lower(), vn)

    def rows_of_type(self, arm_type):
        """Rows of the per-type sheet(s) holding resources of an ARM type (sheet names vary by export)."""
        arm_type = arm_type.lower()
        # newer exports record the exact sheet of every resource; older ones only the friendly type name
        sheets = {str(r.g('Sheet') or r.g('Type')) for r in self.b.sheet('AllResources')
                  if str(r.g('Resource Type') or '').lower() == arm_type}
        names = {str(r['Name']).lower() for r in self.b.sheet('AllResources') if str(r.g('Resource Type') or '').lower() == arm_type}
        # Excel caps sheet names at 31 characters ('Data Collection Rule Associations' -> '...Associatio')
        sheets = {sh if sh in self.b.names else sh[:31] for sh in sheets}
        return [r for s in sheets for r in self.b.sheet(s) if str(r['Name']).lower() in names]

    # disk encryption: disk name -> (DES name or None, encryption type, state)
    def _disks(self):
        self.disk = {}
        for r in self.b.sheet('Disks'):
            des = re.findall(r'diskEncryptionSets/([^/\'"\s]+)', str(r.g('Disk Encryption Set Id') or '') + ' ' + r.text, re.I)
            self.disk[str(r['Name']).lower()] = (des[0] if des else None, str(r.g('Encryption Type') or ''),
                                                 str(r.g('Disk State') or ''))
            if str(r.g('Disk State') or '').lower() == 'unattached':
                sid = self.sub_of_row(r)
                self.stale[sid].append(('compute/Disks', r['Name'], 'unattached disk'))

    def vm_des(self, r):
        """Disk encryption of the VM's managed disks: 'CMK: <DES>' when every disk uses a customer-managed key,
        'CMK on a/b disks' when only some do, else 'platform key'. A disk counts as CMK when its encryption type is
        *CustomerKey* (EncryptionAtRestWithCustomerKey / ...PlatformAndCustomerKeys) or it names a disk encryption
        set - some exports drop the DES id column and only keep the encryption type."""
        names = re.findall(r'/disks/([^/\'"\s,}]+)', str(r.g('Os Disk') or '') + ' ' + str(r.g('Data Disks') or ''), re.I)
        disks = [self.disk[n.lower()] for n in names if n.lower() in self.disk]
        if not disks:
            return ''
        cmk = [d for d in disks if d[0] or 'customerkey' in d[1].lower()]
        sets = sorted({d[0] for d in cmk if d[0]})
        if len(cmk) == len(disks):
            return f"<br><font color='#2E7D32'>CMK: {', '.join(sets) if sets else 'customer-managed key'}</font>"
        if cmk:
            return f"<br><font color='#B9770E'>CMK on {len(cmk)}/{len(disks)} disks</font>"
        return "<br><font color='#9E9E9E'>platform key (no CMK)</font>"

    def _place_resources(self):
        b = self.b
        nic_subnet, nic_primary = {}, {}
        for r in b.sheet('Network Interfaces'):
            s = _subnets_of(r)
            if s:
                nic_subnet[r['Name'].lower()] = s[0]
            nic_primary[r['Name'].lower()] = str(r.g('Primary')).lower() != 'false'
        self.nic_subnet = nic_subnet

        aks_nodes = collections.defaultdict(list)    # (vnet_l, subnet_l) -> AKS node VM names
        for r in b.sheet('Virtual Machines'):
            vn, sn = r.g('V Net', 'VNet'), r.g('Subnet')
            if vn and sn and str(r['Name']).lower().startswith('aks-'):
                aks_nodes[(vn.lower(), str(sn).lower())].append(r['Name'])
                self.placed.add(('Virtual Machines', r['Name'].lower()))
            elif vn and sn:
                self.put(vn, sn, f"{r['Name']}<br>{r.g('Private IP') or ''}{self.vm_des(r)}", 'compute/Virtual_Machine',
                         'Virtual Machines', r['Name'])
            # secondary NICs that sit in other subnets (e.g. firewall trust / untrust legs)
            for nic in re.findall(r"networkInterfaces/([^'/\"]+)", str(r.g('Network Interfaces') or '')):
                s = nic_subnet.get(nic.lower())
                if s and not nic_primary.get(nic.lower(), True) and (not vn or s != (vn.lower(), str(sn).lower())):
                    self.put(*s, f"{r['Name']} (secondary NIC)", 'networking/Network_Interfaces')
        # AKS node VMs (availability-set node pools) collapse into one icon per subnet; names in the tooltip
        self.tooltips = {}
        for (vn, sn), names in aks_nodes.items():
            label = f'AKS node VMs x{len(names)}'
            self.put(vn, sn, label, 'compute/Virtual_Machine')
            self.tooltips[(vn, sn, label)] = '\n'.join(sorted(names))

        for r in b.sheet('Private Endpoints'):
            t = re.findall(r"privateLinkServiceId': '[^']*/providers/([^/]+/[^/]+)/([^/']+)", r.text)
            if not t:
                continue
            typ, target = t[0][0].lower(), t[0][1]
            if typ == 'microsoft.containerservice/managedclusters':
                self.private_aks.add(target.lower())
                self.placed.add(('Private Endpoints', r['Name'].lower()))
                continue
            ic = PE_ICON.get(typ, 'networking/Private_Endpoint')
            if typ == 'microsoft.web/sites' and 'func' not in target.lower():
                ic = 'containers/App_Services'
            # the PE row names its NIC exactly; NIC names are often a truncated form of the PE name
            nic = [nic_subnet[n.lower()] for n in re.findall(r"networkInterfaces/([^'/\"]+)", str(r.g('Network Interfaces')))
                   if n.lower() in nic_subnet]
            s = nic[0] if nic else (_subnets_of(r)[0] if _subnets_of(r) else None)
            if s:
                self.put(*s, f'{target} (PE)', ic, 'Private Endpoints', r['Name'])
                self.pe_targets.add(target.lower())
                self.res_vnet[('pe-target', target.lower())] = self.vlow.get(s[0], s[0])
            else:
                self.unplaced_pe.append(f"{r['Name']} -> {target}")

        vmss_by_subnet = collections.defaultdict(list)
        for r in b.sheet('Virtual Machine Scale Sets'):
            for s in _subnets_of(r)[:1]:
                vmss_by_subnet[s].append(r['Name'])
        aks_subnets = set()
        for r in b.sheet('AKS'):
            ss = _subnets_of(r)
            if not ss:
                continue
            aks_subnets.update(ss)
            priv = ' · private API' if r['Name'].lower() in self.private_aks else ''
            pools = sum(len(vmss_by_subnet[s]) for s in ss)
            self.put(*ss[0], f"{r['Name']}<br>(AKS{priv}, {pools} node pools)", 'containers/Kubernetes_Services', 'AKS', r['Name'])
            for s in ss[1:]:
                for p in vmss_by_subnet[s]:
                    self.put(*s, f"{p}<br>(node pool of {r['Name']})", 'compute/VM_Scale_Sets')
        for s, pools in vmss_by_subnet.items():
            if s not in aks_subnets:
                for p in pools:
                    self.put(*s, p, 'compute/VM_Scale_Sets')

        pip_ip = {r['Name']: r.g('Ip Address') for r in b.sheet('Public IP Addresses')}
        nic_vm = {nic.lower(): r['Name'] for r in b.sheet('Virtual Machines')
                  for nic in re.findall(r"networkInterfaces/([^'/\"]+)", str(r.g('Network Interfaces') or ''))}
        for r in b.sheet('Load Balancers'):
            ss = _subnets_of(r)
            name = r['Name']
            if name == 'kubernetes':
                continue                                  # AKS public LB: counted in the strip
            info = _lb_summary(r, pip_ip, nic_vm)
            if ss:
                label = 'kubernetes-internal (ILB)' if name == 'kubernetes-internal' else name
                self.put(*ss[0], f"{label}<br>{info['short']}", 'networking/Load_Balancers', 'Load Balancers', name)
                self.tooltips[(ss[0][0], ss[0][1], f"{label}<br>{info['short']}")] = info['tip']
            else:
                # public frontend: draw it where its backend NICs are
                back = [nic_subnet[n.lower()] for n in re.findall(r"networkInterfaces/([^'/\"]+)", r.text) if n.lower() in nic_subnet]
                if back:
                    label = f"{name}<br>{info['short']}"
                    self.put(*back[0], label, 'networking/Load_Balancers', 'Load Balancers', name)
                    self.tooltips[(back[0][0], back[0][1], label)] = info['tip']

        simple = [  # (sheet, icon, label suffix, all subnets?)
            ('Application Gateways', 'networking/Application_Gateways', '', False),
            ('Firewalls', 'networking/Firewalls', '<br>(firewall)', True),
            ('Azure Firewalls', 'networking/Firewalls', '<br>(Azure Firewall)', False),
            ('Bastion Hosts', 'networking/Bastions', '', False),
            ('Virtual Network Gateways', 'networking/Virtual_Network_Gateways', '<br>(gateway)', False),
            ('Nat Gateways', 'networking/NAT', '<br>(NAT gateway - outbound)', False),
            ('API Management', 'integration/API_Management_Services', '', False),
            ('Container Groups', 'containers/Container_Instances', '', False),
            ('SQL Managed Instances', 'databases/SQL_Managed_Instance', '', False),
        ]
        # NAT gateway <-> the subnets whose outbound internet traffic it carries (marker on each subnet)
        self.subnet_nat = {}
        for r in b.sheet('Nat Gateways'):
            pips = [pip_ip.get(p) or p for p in re.findall(r"publicIPAddresses/([^'/\"]+)", str(r.g('Public Ip Addresses') or ''))]
            for s in _subnets_of(r):
                self.subnet_nat[s] = (r['Name'], pips)
        for sheet, ic, suffix, every in simple:
            for r in b.sheet(sheet):
                for s in (_subnets_of(r) if every else _subnets_of(r)[:1]):
                    self.put(*s, f"{r['Name']}{suffix}", ic, sheet, r['Name'])
        for sheet, ic in (('PostgreSQL Flexible Servers', 'databases/Azure_Database_PostgreSQL_Server'),
                          ('MySQL Flexible Servers', 'databases/Azure_Database_MySQL_Server')):
            for r in b.sheet(sheet):
                for s in _subnets_of(r)[:1]:
                    self.put(*s, f"{r['Name']}<br>(VNet-integrated)", ic, sheet, r['Name'])
        for r in b.sheet('App Service'):
            for s in _subnets_of(r)[:1]:
                ic = 'compute/Function_Apps' if 'func' in r['Name'].lower() else 'containers/App_Services'
                self.put(*s, f"{r['Name']}<br>(VNet integration)", ic, 'App Service', r['Name'])
        n_apps = len(self.rows_of_type('microsoft.app/containerapps') or b.sheet('Container Apps'))
        n_jobs = len(self.rows_of_type('microsoft.app/jobs') or b.sheet('Jobs'))
        for r in b.sheet('Managed Environments'):
            for s in _subnets_of(r)[:1]:
                self.put(*s, f"{r['Name']}<br>(Container Apps env: {n_apps} apps, {n_jobs} jobs)",
                         'other/Container_App_Environments', 'Managed Environments', r['Name'])
        for sheet, kind in (('Inbound Endpoints', 'inbound'), ('Outbound Endpoints', 'outbound')):
            for r in b.sheet(sheet):
                ip = re.findall(r"privateIpAddress': '([\d.]+)", r.text)
                for s in _subnets_of(r)[:1]:
                    self.put(*s, f"{r['Name']}<br>DNS resolver {kind}{' ' + ip[0] if ip else ''}", 'networking/DNS_Private_Resolver')

    # public IPs and NSGs: small markers on what they are attached to; anything attached to nothing is stale
    def _attachments(self):
        b = self.b
        nic_vm = {}
        for r in b.sheet('Virtual Machines'):
            for nic in re.findall(r"networkInterfaces/([^'/\"]+)", str(r.g('Network Interfaces') or '')):
                nic_vm[nic.lower()] = r['Name']
        pe_nics = {n.lower() for r in b.sheet('Private Endpoints')
                   for n in re.findall(r"networkInterfaces/([^'/\"]+)", str(r.g('Network Interfaces') or ''))}

        for r in b.sheet('Public IP Addresses'):
            ip = r.g('Ip Address') or 'no IP'
            tip = f"Public IP {r['Name']} ({ip})"
            m = re.match(r'\s*([^:]+):\s*(.+)', str(r.g('Attached To') or ''))
            if not m or str(r.g('Attached')).lower() == 'no':
                self.stale[self.sub_of_row(r)].append(('networking/Public_IP_Addresses', r['Name'], f'public IP {ip} - not attached'))
                continue
            kind, owner = m.group(1).strip(), m.group(2).strip()
            if kind == 'Network Interfaces':
                vm = nic_vm.get(owner.lower())
                if not vm:
                    self.stale[self.sub_of_row(r)].append(
                        ('networking/Public_IP_Addresses', r['Name'], f'public IP {ip} on NIC {owner} with no VM'))
                    continue
                kind, owner = 'Virtual Machines', vm
            self.markers[(kind, owner.lower())].append((_icon('networking/Public_IP_Addresses'), tip))

        for r in b.sheet('Network Interfaces'):
            n = r['Name'].lower()
            if n not in nic_vm and n not in pe_nics and not re.search(r'(?i)\.nic\.|vmss|aks-', r['Name']):
                self.stale[self.sub_of_row(r)].append(('networking/Network_Interfaces', r['Name'], 'NIC not attached to a VM'))

        # NSG on subnets: the VNet sheet lists them aligned with the subnet names; the NSG sheet lists its subnets
        for r in b.sheet('Virtual Networks'):
            names = [s.strip() for s in str(r.g('Subnet Names') or '').split(';')]
            nsgs = [s.strip() for s in str(r.g('Subnet NSG IDs') or '').split(';')]
            for s, nid in zip(names, nsgs):
                m = re.search(r'networkSecurityGroups/([^/\'"\s]+)', nid, re.I)
                if s and m:
                    self.subnet_nsg[(r['Name'].lower(), s.lower())] = m.group(1)
        for r in b.sheet('Network Security Groups'):
            subs = _subnets_of(r)
            nics = re.findall(r"networkInterfaces/([^'/\"]+)", r.text)
            for s in subs:
                self.subnet_nsg.setdefault(s, r['Name'])
            for nic in nics:
                vm = nic_vm.get(nic.lower())
                if vm:
                    self.markers[('Virtual Machines', vm.lower())].append((_icon('networking/Network_Security_Groups'),
                                                                          f"NIC-level NSG {r['Name']}"))
            if not subs and not nics and r['Name'].lower() not in {v.lower() for v in self.subnet_nsg.values()}:
                self.stale[self.sub_of_row(r)].append(('networking/Network_Security_Groups', r['Name'], 'NSG not attached'))

    # route tables: what each subnet's traffic hops to next, with the next-hop IP resolved to the resource behind it
    def _routes(self):
        b = self.b
        nic_vm = {nic.lower(): r['Name'] for r in b.sheet('Virtual Machines')
                  for nic in re.findall(r"networkInterfaces/([^'/\"]+)", str(r.g('Network Interfaces') or ''))}
        ip_owner = {}
        for r in b.sheet('Network Interfaces'):
            for ip in re.findall(r"privateIPAddress': '([\d.]+)", r.text):
                vm = nic_vm.get(r['Name'].lower())
                ip_owner.setdefault(ip, f"VM {vm}" if vm else f"NIC {r['Name']}")
        for sheet, kind in (('Load Balancers', 'LB'), ('Firewalls', 'firewall'), ('Azure Firewalls', 'Azure Firewall'),
                            ('Application Gateways', 'App Gateway'), ('Virtual Network Gateways', 'gateway')):
            for r in b.sheet(sheet):
                for ip in re.findall(r"privateIPAddress': '([\d.]+)", r.text):
                    ip_owner[ip] = f"{kind} {r['Name']}"
        self.subnet_rt = {}
        for r in self.rows_of_type('microsoft.network/routetables') or b.sheet('Route Tables'):
            text = str(r.g('Routes') or '[]')
            summary = str(r.g('Routes Summary') or '')
            truncated = False
            if summary:
                # the collector's compact form, "name: prefix -> NextHopType [ip]; ...", stays under Excel's cell limit
                raw = []
                for part in summary.split('; '):
                    m = re.match(r'(.*?): (\S+) -> (\S+)(?: (\S+))?$', part.strip())
                    if m:
                        raw.append((m.group(1), {'addressPrefix': m.group(2), 'nextHopType': m.group(3),
                                                 'nextHopIpAddress': m.group(4) or ''}))
            else:
                try:
                    raw = [(x.get('name', ''), x.get('properties', x)) for x in ast.literal_eval(text)]
                except (ValueError, SyntaxError):
                    # Excel caps a cell at 32,767 characters, so long route lists arrive cut off: read every complete
                    # route's properties block, whatever order its fields are in
                    truncated = True
                    raw = []
                    for n, body in re.findall(r"'name': '([^']*)', 'properties': \{([^{}]*)\}", text):
                        f = dict(re.findall(r"'(addressPrefix|nextHopType|nextHopIpAddress)': '([^']*)'", body))
                        if f.get('addressPrefix') and f.get('nextHopType'):
                            raw.append((n, f))
            routes = []
            for name, pr in raw:
                hop, ip = pr.get('nextHopType', ''), pr.get('nextHopIpAddress') or ''
                routes.append(dict(name=name, prefix=pr.get('addressPrefix', ''), hop=hop, ip=ip,
                                   owner=ip_owner.get(ip, '') if ip else ''))
            subs = [s for s in _subnets_of(r)]
            if not subs:
                self.stale[self.sub_of_row(r)].append(('networking/Route_Tables', r['Name'], 'route table not attached to a subnet'))
            for s in subs:
                self.subnet_rt[s] = dict(name=r['Name'], routes=routes, truncated=truncated,
                                         bgp_off=str(r.g('Disable Bgp Route Propagation')).lower() == 'true')

    @staticmethod
    def route_text(rt):
        """One-line summary (the default route when there is one) and the full tooltip of a subnet's route table."""
        def hop(x):
            if x['hop'] == 'VirtualAppliance':
                return f"NVA {x['ip']}" + (f" ({x['owner']})" if x['owner'] else '')
            return {'Internet': 'Internet', 'VnetLocal': 'VNet local', 'None': 'dropped (None)',
                    'VirtualNetworkGateway': 'VPN/ER gateway'}.get(x['hop'], x['hop'] or '?')
        routes = rt['routes']
        main = next((x for x in routes if x['prefix'] == '0.0.0.0/0'), routes[0] if routes else None)
        short = f"UDR {rt['name']}: " + (f"{main['prefix']} → {hop(main)}" if main else 'no routes')
        if len(routes) > 1:
            short += f" +{len(routes) - 1}{'+' if rt.get('truncated') else ''}"
        tip = (f"Route table {rt['name']}" + (' (BGP route propagation disabled)' if rt['bgp_off'] else '') +
               (f"\nOnly the first {len(routes)} routes are in the workbook - the list hit Excel's 32,767-character cell limit"
                if rt.get('truncated') else '') + '\n' +
               '\n'.join(f"  {x['name']}: {x['prefix']} -> {hop(x)}" for x in routes))
        return short, tip

    # hybrid connectivity: ExpressRoute circuits, on-prem sites (local network gateways) and the connections joining
    # them to VNet gateways. The collector lifts the connection's gateway / site / circuit names into plain columns
    # (the export drops nested ids); without them the connections are listed unlinked.
    def _hybrid(self):
        circuits = [dict(name=r['Name'], detail=' · '.join(str(x) for x in (r.g('Service Provider Name'), (str(r.g('Bandwidth In Mbps')) + ' Mbps')
                                                                             if r.g('Bandwidth In Mbps') else None,
                                                                             r.g('Peering Location')) if x),
                         state=str(r.g('Service Provider Provisioning State') or ''))
                    for r in self.rows_of_type('microsoft.network/expressroutecircuits')]
        sites = []
        for r in self.rows_of_type('microsoft.network/localnetworkgateways'):
            pref = [str(x) for x in _as_list(r.g('Address Prefixes'))]
            sites.append(dict(name=r['Name'], detail=f"on-prem {r.g('Gateway Ip Address') or r.g('Fqdn') or ''}"
                                                     + (f" · {', '.join(pref[:3])}{' +' + str(len(pref) - 3) if len(pref) > 3 else ''}" if pref else '')))
        conns = []
        for r in self.rows_of_type('microsoft.network/connections'):
            def link(col, kind):
                v = r.g(col)
                if v:
                    return str(v)
                m = re.findall(kind + r'/([^/\'"\s]+)', r.text, re.I)
                return m[0] if m else None
            gw = link('Virtual Network Gateway1 Name', 'virtualNetworkGateways')
            remote = (link('Local Network Gateway2 Name', 'localNetworkGateways') or link('Peer Name', 'expressRouteCircuits')
                      or link('Virtual Network Gateway2 Name', 'virtualNetworkGateways/__never__'))
            conns.append(dict(name=r['Name'], type=str(r.g('Connection Type') or ''), status=str(r.g('Connection Status') or ''),
                              gw=gw, remote=remote,
                              vnet=self.res_vnet.get(('res', 'Virtual Network Gateways', str(gw).lower())) if gw else None))
        self.hybrid = dict(circuits=circuits, sites=sites, conns=conns)

    # private DNS zones linked to each VNet (the collector lifts zone + VNet names into plain columns)
    def _dns_links(self):
        self.vnet_dns = collections.defaultdict(list)
        for r in self.rows_of_type('microsoft.network/privatednszones/virtualnetworklinks') or self.b.sheet('Virtual Network Links'):
            zone = r.g('Private Dns Zone Name') or (re.findall(r'privateDnsZones/([^/\'"\s]+)', r.text, re.I) or [None])[0]
            vnet = r.g('Virtual Network Name') or (re.findall(r'virtualNetworks/([^/\'"\s]+)', r.text, re.I) or [None])[0]
            if zone and vnet:
                self.vnet_dns[str(vnet).lower()].append((str(zone), str(r.g('Registration Enabled')).lower() == 'true'))

    # governance: Defender plans, policy assignments, resource locks, backup-protected items
    def _governance(self):
        self.gov = collections.defaultdict(lambda: dict(defender=[], policy=[], locks=[], backup=[]))
        self.backed_up = set()
        for r in self.rows_of_type('microsoft.security/pricings'):
            self.gov[self.sub_of_row(r)]['defender'].append((r['Name'], str(r.g('Pricing Tier') or '')))
        for r in self.rows_of_type('microsoft.authorization/policyassignments'):
            self.gov[self.sub_of_row(r)]['policy'].append(dict(
                name=str(r.g('Display Name') or r['Name']), mode=str(r.g('Enforcement Mode') or ''),
                noncompliant=int(float(r.g('Non Compliant Resources') or 0)), evaluated=int(float(r.g('Evaluated Resources') or 0))))
        for r in self.rows_of_type('microsoft.authorization/locks'):
            target = str(r.g('Locked Resource Name') or '')
            self.gov[self.sub_of_row(r)]['locks'].append(dict(level=str(r.g('Level') or ''), target=target,
                                                              ttype=str(r.g('Locked Resource Type') or '')))
            if target and '/' in str(r.g('Locked Resource Type') or ''):
                self.markers_lock.setdefault(target.lower(), str(r.g('Level') or 'lock'))
        for r in self.rows_of_type('microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems'):
            name = str(r.g('Protected Resource Name') or r['Name'])
            status = str(r.g('Last Backup Status') or '')
            self.gov[self.sub_of_row(r)]['backup'].append(dict(name=name, status=status, state=str(r.g('Protection State') or ''),
                                                               vault=str(r.g('Vault Name') or ''), when=str(r.g('Last Backup Time') or '')))
            self.backed_up.add(name.lower())
            self.markers[('Virtual Machines', name.lower())].append(
                (_icon('other/Backup_Vault'), f"Backed up to {r.g('Vault Name') or 'vault'} - last backup {status or 'n/a'}"
                                              f"{' at ' + str(r.g('Last Backup Time')) if r.g('Last Backup Time') else ''}"))
        self.has_backup_data = bool(self.rows_of_type('microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems')) \
            or any('Backup Protected Items' in n for n in self.b.names)
        for target, level in self.markers_lock.items():
            for kind in ('Virtual Machines', 'Load Balancers', 'Application Gateways', 'Virtual Network Gateways', 'Firewalls'):
                self.markers[(kind, target)].append((_icon('management_governance/Customer_Lockbox_for_MS_Azure'), f'Resource lock: {level}'))

    # what each Log Analytics workspace receives from, and what each user-assigned identity is assigned to
    def _links(self):
        b = self.b
        skip = {'AllResources', 'IAM Roles', 'Dashboard', 'RG', 'Workspaces', 'User Assigned Identities', 'Subscriptions', 'Truncated Cells',
                'Scheduled Query Rules', 'Activity Log Alerts', 'Metric Alerts', 'Action Group', 'Smart Detector Alert Rules'}
        ws_rows = self.rows_of_type('microsoft.operationalinsights/workspaces') or b.sheet('Workspaces')
        ws_names = {str(r['Name']).lower(): r['Name'] for r in ws_rows}
        mi_rows = b.sheet('User Assigned Identities')
        mi_names = {str(r['Name']).lower(): r['Name'] for r in mi_rows}
        self.law_links = collections.defaultdict(set)
        self.mi_links = collections.defaultdict(set)
        for sheet in b.names:
            if sheet in skip or sheet.startswith(('High-Level', 'Low-Level')) or sheet.endswith(('_Quota', '_Quot', 'Quota')):
                continue
            for r in b.sheet(sheet):
                if not r.get('Name'):
                    continue
                who = f"{r['Name']} ({sheet})"
                for w in re.findall(r'operationalinsights/workspaces/([^/\'"\s,}\]]+)', r.text, re.I):
                    if w.lower() in ws_names:
                        self.law_links[w.lower()].add(who)
                # identities appear as resource ids in add-on blocks, and as column names
                # 'Managed Identity/User Assigned Identities/<Name>' carrying a value when assigned
                for m in re.findall(r'userAssignedIdentities/([^/\'"\s,}\]]+)', r.text, re.I):
                    if m.lower() in mi_names:
                        self.mi_links[m.lower()].add(who)
                for k, v in r.items():
                    k = str(k or '')
                    if 'user assigned identities/' in k.lower() and v not in (None, '', 'None'):
                        m = k.split('/')[-1].strip().lower()
                        if m in mi_names:
                            self.mi_links[m].add(who)
        # data collection rules: what they collect, which workspace(s) they send to, which resources they cover
        # (associations come from the collector's 'Data Collection Rule Associations' sheet when present)
        assoc = collections.defaultdict(list)
        for a in self.rows_of_type('microsoft.insights/datacollectionruleassociations'):
            dcr = str(a.g('Data Collection Rule Name') or '')
            if dcr:
                assoc[dcr.lower()].append(f"{a.g('Associated Resource Name') or a['Name']}"
                                          f" ({a.g('Associated Resource Type') or 'resource'})")
        kinds = {'Windows Event Logs': 'Windows events', 'Syslog': 'syslog', 'Performance Counters': 'perf counters',
                 'Log Files': 'log files', 'Extensions': 'extensions', 'Prometheus Forwarder': 'Prometheus', 'Iis Logs': 'IIS logs'}
        self.law_dcrs = collections.defaultdict(list)      # workspace name (lower) -> [dcr dict]
        self.has_assoc = bool(assoc)
        for r in self.b.sheet('Data Collection Rules'):
            dests = {w.lower() for w in re.findall(r'operationalinsights/workspaces/([^/\'"\s,}\]]+)', r.text, re.I)}
            collects = [label for col, label in kinds.items() if r.g(col) not in (None, '', 'None', '[]')]
            dce = re.findall(r'dataCollectionEndpoints/([^/\'"\s]+)', str(r.g('Data Collection Endpoint Id') or ''), re.I)
            dcr = dict(name=r['Name'], collects=collects, resources=sorted(set(assoc.get(str(r['Name']).lower(), []))),
                       dce=dce[0] if dce else None)
            for w in dests:
                if w in ws_names:
                    self.law_dcrs[w].append(dcr)
        # diagnostic settings (collected per resource): resource -> workspace / storage / event hub
        for r in self.rows_of_type('microsoft.insights/diagnosticsettings'):
            ws = str(r.g('Workspace Name') or '').lower()
            if ws in ws_names:
                self.law_links[ws].add(f"{r.g('Target Resource Name') or r['Name']} (diagnostic setting)")
        # Application Insights: component -> workspace, and the apps it monitors. Apps link through the collector's
        # 'App Insights Component Name' column, else the hidden-link tags Azure writes on the app (resource id, or
        # the connection string / instrumentation key matched to the component's key).
        comps = self.rows_of_type('microsoft.insights/components') or b.sheet('Application Insights')
        by_key = {}
        for c in comps:
            for k in (c.g('Instrumentation Key'), c.g('App Id'), c.g('Application Id')):
                if k and '***' not in str(k):
                    by_key[str(k).lower()] = c['Name']
        comp_names = {str(c['Name']).lower(): c['Name'] for c in comps}
        apps_of = collections.defaultdict(list)
        for r in self.rows_of_type('microsoft.web/sites') or b.sheet('App Service'):
            comp = r.g('App Insights Component Name')
            if not comp:
                m = re.search(r'microsoft\.insights/components/([^"\'/\s]+)', r.text, re.I)
                comp = comp_names.get(m.group(1).lower()) if m else None
            if not comp:
                for k in re.findall(r'InstrumentationKey=([0-9a-f-]{36})|instrumentation-key[^0-9a-f]{0,12}([0-9a-f-]{36})', r.text, re.I):
                    comp = by_key.get((k[0] or k[1]).lower())
                    if comp:
                        break
            if comp:
                apps_of[str(comp).lower()].append(r['Name'])
                self.markers[('App Service', str(r['Name']).lower())].append(
                    (_icon('devops/Application_Insights'), f'Monitored by Application Insights {comp}'))
        self.law_ai = collections.defaultdict(list)       # workspace (lower) -> [component]
        self.ai_classic = collections.defaultdict(list)   # subscription -> [component without a workspace]
        for c in comps:
            ws = (re.findall(r'workspaces/([^/\'"\s]+)', str(c.g('Workspace Resource Id') or ''), re.I)
                  or [c.g('Workspace Resource Name')])[0]
            ai = dict(name=c['Name'], apps=sorted(apps_of.get(str(c['Name']).lower(), [])),
                      kind=str(c.g('Application Type', 'Application_ Type', 'Kind') or 'web'),
                      retention=str(c.g('Retention In Days') or ''), sampling=str(c.g('Sampling Percentage') or ''))
            if ws and str(ws).lower() in ws_names:
                self.law_ai[str(ws).lower()].append(ai)
            else:
                self.ai_classic[self.sub_of_row(c)].append(ai)
        self.law_cards = collections.defaultdict(list)    # subscription id -> [(name, sorted links)]
        for r in ws_rows:
            self.law_cards[self.sub_of_row(r)].append((r['Name'], sorted(x for x in self.law_links[str(r['Name']).lower()]
                                                                         if not x.endswith(('(Data Collection Rules)',
                                                                                            '(Application Insights)')))))
        self.mi_cards = collections.defaultdict(list)
        for r in mi_rows:
            links = sorted(self.mi_links[str(r['Name']).lower()])
            if links:
                self.mi_cards[self.sub_of_row(r)].append((r['Name'], links))
            else:
                self.stale[self.sub_of_row(r)].append(('identity/Managed_Identities', r['Name'], 'identity not assigned to any resource'))

    # Front Door origins (collected by the assessment script from the Front Door API) -> the backend they reach
    def _frontdoor(self):
        """self.afd: [(profile row, [origin])], origin = dict(what, host, group, icon, vnet). vnet is set when the
        origin resolves to something drawn inside a VNet (public IP owner, private link target, PE'd App Service /
        storage), so an arrow can be routed to it; other origins are drawn in the box under the profile."""
        self.afd = []
        self.linked_wafs = set()
        profiles = self.rows_of_type('microsoft.cdn/profiles')
        origins = self.rows_of_type('microsoft.cdn/profiles/origingroups/origins')
        ip_owner = {}
        for r in self.b.sheet('Public IP Addresses'):
            m = re.match(r'\s*([^:]+):\s*(.+)', str(r.g('Attached To') or ''))
            if m and r.g('Ip Address'):
                ip_owner[str(r.g('Ip Address'))] = (m.group(1).strip(), m.group(2).strip())
        nic_vm = {nic.lower(): r['Name'] for r in self.b.sheet('Virtual Machines')
                  for nic in re.findall(r"networkInterfaces/([^'/\"]+)", str(r.g('Network Interfaces') or ''))}
        by_type = {'Load Balancers', 'Application Gateways', 'Virtual Machines', 'App Service', 'Firewalls'}
        rg_count = collections.Counter(str(p.g('Resource Group') or '').lower() for p in profiles)

        def belongs(o, p):
            # the export drops nested '.id' columns, so use the collector's plain 'Profile Name' column when present,
            # then the id path, then the resource group (when it holds a single profile), then the only profile
            pn = o.g('Profile Name')
            if pn:
                return str(pn).lower() == str(p['Name']).lower()
            if '/profiles/' in o.text.lower():
                return f"/profiles/{p['Name']}/".lower() in o.text.lower()
            rg = str(o.g('Resource Group') or '').lower()
            if rg and rg_count.get(rg) == 1:
                return rg == str(p.g('Resource Group') or '').lower()
            return len(profiles) == 1

        def resolve(host, link_text, name, group):
            """Origin -> dict(what, host, group, icon, vnet): the backend resource it reaches."""
            item = dict(what=host or name, host=host, group=group, icon='networking/DNS_Zones', vnet=None, name=name)
            link = re.findall(r'/providers/([^/]+/[^/]+)/([^/\'"\s]+)', link_text)
            if link:
                tname = link[0][1]
                item.update(what=f'{tname} (private link)', icon=PE_ICON.get(link[0][0].lower(), 'networking/Private_Link'),
                            vnet=(self.res_vnet.get(('pe-target', tname.lower()))
                                  or next((self.res_vnet[('res', t, tname.lower())] for t in by_type
                                           if ('res', t, tname.lower()) in self.res_vnet), None)))
            elif host in ip_owner:
                kind, owner = ip_owner[host]
                if kind == 'Network Interfaces':
                    kind, owner = 'Virtual Machines', nic_vm.get(owner.lower(), owner)
                item.update(what=f'{owner} ({host})', icon='networking/Load_Balancers' if kind == 'Load Balancers'
                            else 'compute/Virtual_Machine' if kind == 'Virtual Machines' else 'networking/Public_IP_Addresses',
                            vnet=self.res_vnet.get(('res', kind, owner.lower())))
            else:
                # PaaS host names: <app>.azurewebsites.net, <account>[.zNN].(web|blob).core.windows.net
                m = re.match(r'([^.]+)\.(?:z\d+\.)?(azurewebsites\.net|blob\.core\.windows\.net|web\.core\.windows\.net)$', host, re.I)
                if m:
                    app = 'websites' in m.group(2).lower()
                    kind = 'App Service' if app else ('Storage static website' if m.group(2).lower().startswith('web') else 'Storage blob')
                    item.update(what=f'{m.group(1)} ({kind})', icon='containers/App_Services' if app else 'storage/Storage_Accounts',
                                vnet=self.res_vnet.get(('pe-target', m.group(1).lower()))
                                or self.res_vnet.get(('res', 'App Service', m.group(1).lower())))
            return item

        def as_list(v):
            try:
                x = ast.literal_eval(str(v)) if v not in (None, '', 'None') else []
                return x if isinstance(x, list) else []
            except (ValueError, SyntaxError):
                return []

        for p in profiles:
            items = []
            for o in origins:
                if not belongs(o, p):
                    continue
                link_text = ' '.join(str(v) for k, v in o.items() if 'privatelink' in _norm(k))
                items.append(resolve(str(o.g('Host Name', 'Hostname') or ''), link_text, o['Name'],
                                     str(o.g('Origin Group Name', 'Origin Group') or '')))

            # the way in: endpoint host names (*.azurefd.net) and custom domains, each with the routes that serve it
            routes = [r for r in self.rows_of_type('microsoft.cdn/profiles/afdendpoints/routes') if belongs(r, p)]

            def served_by(match):
                return [f"{r['Name']}: {r.g('Patterns To Match') or '/*'} -> {r.g('Origin Group Name') or '?'}"
                        for r in routes if match(r)]

            domains = []
            for c in self.rows_of_type('microsoft.cdn/profiles/customdomains'):
                if belongs(c, p):
                    host = str(c.g('Host Name') or c['Name'])
                    domains.append(dict(host=host, kind='custom domain', state=str(c.g('Domain Validation State') or ''),
                                        routes=served_by(lambda r, n=c['Name']: n.lower() in str(r.g('Custom Domain Names') or '').lower())))
            for e in self.rows_of_type('microsoft.cdn/profiles/afdendpoints'):
                if belongs(e, p):
                    domains.append(dict(host=str(e.g('Host Name') or e['Name']), kind='endpoint', state=str(e.g('Enabled State') or ''),
                                        routes=served_by(lambda r, n=e['Name']: str(r.g('Endpoint Name') or '').lower() == n.lower())))

            # classic Azure CDN / Front Door classic: each endpoint carries its own custom domains and origins
            # (Resource Graph returns them inline), so the endpoint itself is the route
            classic = [e for e in self.rows_of_type('microsoft.cdn/profiles/endpoints') if belongs(e, p)]
            for e in classic:
                state = str(e.g('Resource State') or '')
                e_origins = [resolve(str((o.get('properties') or {}).get('hostName') or ''), '', o.get('name', ''),
                                     f"endpoint {e['Name']}") for o in as_list(e.g('Origins'))]
                items += e_origins
                via = [f"endpoint {e['Name']}{' (' + state + ')' if state else ''} -> " + ', '.join(o['what'] for o in e_origins)]
                for cd in as_list(e.g('Custom Domains')):
                    host = (cd.get('properties') or {}).get('hostName') or cd.get('name')
                    domains.append(dict(host=host, kind=f"custom domain · {e['Name']}", state=state, routes=via))
                domains.append(dict(host=str(e.g('Host Name') or e['Name']),
                                    kind='endpoint' + (' · STOPPED' if state.lower() == 'stopped' else ''), state=state, routes=via))
            # one row per route: the domains it serves -> the route -> the origin backends of its origin group
            cd_rows = [c for c in self.rows_of_type('microsoft.cdn/profiles/customdomains') if belongs(c, p)]
            cd_host = {str(c['Name']).lower(): str(c.g('Host Name') or c['Name']) for c in cd_rows}
            cd_state = {str(c['Name']).lower(): str(c.g('Domain Validation State') or '') for c in cd_rows}
            ep_host = {str(e['Name']).lower(): str(e.g('Host Name') or e['Name'])
                       for e in self.rows_of_type('microsoft.cdn/profiles/afdendpoints') if belongs(e, p)}
            okey = lambda o: (o['name'], o['host'])
            rows, used_hosts, used_origins = [], set(), set()
            for r in routes:
                hosts = []
                for n in [x.strip().lower() for x in str(r.g('Custom Domain Names') or '').split(',') if x.strip()]:
                    st = cd_state.get(n, '')
                    hosts.append((cd_host.get(n, n), 'custom domain' + (f' · {st}' if st and st != 'Approved' else '')))
                ep = str(r.g('Endpoint Name') or '').lower()
                if ep in ep_host and str(r.g('Link To Default Domain') or 'Enabled').lower() != 'disabled':
                    hosts.append((ep_host[ep], 'endpoint'))
                og = str(r.g('Origin Group Name') or '').lower()
                backends = [o for o in items if o['group'].lower() == og]
                used_hosts.update(h for h, _ in hosts)
                used_origins.update(okey(o) for o in backends)
                detail = ' · '.join(x for x in (
                    str(r.g('Patterns To Match') or '/*'), str(r.g('Forwarding Protocol') or ''),
                    'HTTPS redirect' if str(r.g('Https Redirect')).lower() == 'enabled' else '',
                    'DISABLED' if str(r.g('Enabled State')).lower() == 'disabled' else '') if x)
                rows.append(dict(domains=hosts, route=dict(name=r['Name'], detail=detail, group=str(r.g('Origin Group Name') or '')),
                                 backends=backends))
            for e in classic:
                state = str(e.g('Resource State') or '')
                hosts = [((cd.get('properties') or {}).get('hostName') or cd.get('name'), 'custom domain')
                         for cd in as_list(e.g('Custom Domains'))]
                hosts.append((str(e.g('Host Name') or e['Name']), 'endpoint'))
                names = {o.get('name') for o in as_list(e.g('Origins'))}
                backends = [o for o in items if o['name'] in names and o['group'] == f"endpoint {e['Name']}"]
                used_hosts.update(h for h, _ in hosts)
                used_origins.update(okey(o) for o in backends)
                rows.append(dict(domains=hosts, route=dict(name=f"endpoint {e['Name']}", detail=state or 'classic endpoint',
                                                           group=''), backends=backends))
            stray_hosts = [(dm['host'], dm['kind']) for dm in domains if dm['host'] not in used_hosts]
            stray_origins = [o for o in items if okey(o) not in used_origins]
            if stray_hosts or stray_origins:
                rows.append(dict(domains=stray_hosts, route=None, backends=stray_origins))
            wafs = []
            for w in self.rows_of_type('microsoft.network/frontdoorwebapplicationfirewallpolicies'):
                if f"/profiles/{str(p['Name']).lower()}/" in w.text.lower():
                    custom = len(re.findall(r"'name': '", str(w.g('Rules') or '')))
                    managed = re.findall(r"'ruleSetType': '([^']+)'", str(w.g('Managed Rule Sets') or ''))
                    wafs.append(dict(name=w['Name'], mode=str(w.g('Policy Settings Mode') or ''),
                                     state=str(w.g('Enabled State') or ''), custom=custom, managed=managed))
                    self.linked_wafs.add(str(w['Name']).lower())
            kind = 'Front Door' if not classic else ('CDN / Front Door (classic)' if not routes else 'Front Door + classic CDN')
            self.afd.append(dict(profile=p, origins=items, domains=domains, routes=len(routes) + len(classic), kind=kind,
                                 classic=bool(classic), rows=rows, wafs=wafs))

    def _strip(self):
        """Per-subscription counts of resources NOT drawn inside a subnet (a storage account or key vault reached
        through a private endpoint is drawn in that subnet and not counted again). Public IPs, NSGs, identities and
        workspaces have their own markers / panels."""
        handled = {'Network Security Groups', 'Public IP Addresses', 'User Assigned Identities', 'Workspaces', 'Route Tables'}
        self.strip = collections.defaultdict(lambda: collections.defaultdict(list))     # sid -> arm type -> names
        self.arm_sheet = {}
        self.sub_locs = collections.defaultdict(collections.Counter)
        for r in self.b.sheet('AllResources'):
            sid = str(r.g('Subscription') or '').lower()
            typ = str(r.g('Type') or '')
            loc = str(r.g('Location') or '').lower()
            if sid not in self.sub_name or not typ:
                continue
            if loc and loc != 'global':
                self.sub_locs[sid][loc] += 1
            name = str(r['Name']).lower()
            arm = str(r.g('Resource Type') or '').lower()
            if (arm in ARM_HANDLED or arm in ARM_SKIP or (typ, name) in self.placed or name in self.pe_targets
                    or (arm == 'microsoft.network/frontdoorwebapplicationfirewallpolicies' and name in self.linked_wafs)):
                continue
            # mapped ARM types are always listed; otherwise fall back to the sheet-name rules
            if arm not in ARM_STRIP and (typ in STRIP_SKIP or typ in handled):
                continue
            self.arm_sheet.setdefault(arm, typ)
            self.strip[sid][arm or typ].append(str(r['Name']))


# ---------------------------------------------------------------- layout + drawing

class _Builder:
    def __init__(self, inv, title_source):
        self.inv = inv
        self.d = _Diagram('Azure landing zone (current state)')
        self.cell_of = {}
        self.source = title_source
        self._roles()

    # ----- roles: hubs per region, transit VNets for badges, home region per subscription
    def _roles(self):
        inv = self.inv
        V = inv.vnets
        self.in_peers = {n: [p for p in v['peers'] if p in V] for n, v in V.items()}
        by_loc = collections.defaultdict(list)
        for n, v in V.items():
            by_loc[v['loc']].append(n)
        self.hub = {}
        for loc, names in by_loc.items():
            # most peerings to inventoried VNets, then most peerings overall (small tenants peer mostly outward)
            best = max(names, key=lambda n: (len(self.in_peers[n]), len(V[n]['peers'])))
            if len(V[best]['peers']) >= 2:
                self.hub[loc] = best
        if self.hub:
            self.primary = max(self.hub, key=lambda l: (len(self.in_peers[self.hub[l]]), len(V[self.hub[l]]['peers'])))
        else:
            self.primary = max(by_loc, key=lambda l: len(by_loc[l])) if by_loc else 'unknown'
        self.regions = [self.primary] + sorted((l for l in by_loc if l != self.primary), key=lambda l: -len(by_loc[l]))

        # external (peered, not inventoried) VNets
        self.outside = collections.defaultdict(set)
        for n, v in V.items():
            for p in v['peers']:
                if p not in V:
                    self.outside[p].add(n)

        # badge for every VNet that 3+ VNets peer with
        cnt = collections.Counter()
        for v in V.values():
            cnt.update(v['peers'])
        transit = [n for n, c in cnt.most_common() if c >= 3][:len(BADGE_COLORS)]
        self.badges = []
        for i, n in enumerate(transit):
            if n == self.hub.get(self.primary):
                label = 'HUB'
            elif n in self.hub.values():
                label = f'HUB {_region_title(V[n]["loc"]).upper()}'
            else:
                label = _short(n)
            self.badges.append((n, label, BADGE_COLORS[i]))

        # firewall VNets: hold a firewall resource or firewall-named subnets
        self.firewall_vnets = set()
        for n, v in V.items():
            items = [l for s in v['subnets'] for l, _ in inv.place[(n.lower(), s.lower())]]
            if (re.search(r'(?i)(^|[-_])(fw|ngfw|firewall)([-_]|$)', n)
                    or any(re.search(r'(?i)firewall|palo|ngfw', s) for s in v['subnets'])
                    or any('firewall' in l.lower() for l in items)):
                if n not in self.hub.values():
                    self.firewall_vnets.add(n)

        # home region of each subscription: where most of its VNets are, else most of its resources
        self.sub_home = {}
        for sid in inv.sub_name:
            locs = collections.Counter(v['loc'] for v in V.values() if v['sub'] == sid)
            if not locs:
                locs = inv.sub_locs.get(sid) or collections.Counter({self.primary: 1})
            home = locs.most_common(1)[0][0]
            self.sub_home[sid] = home if home in self.regions else self.primary

    # ----- VNet box
    def vnet_size(self, v):
        return self._vnet(None, v, 0, 0, measure=True)

    def _vnet(self, parent, v, x, y, measure=False):
        inv, d = self.inv, self.d
        subs = v['subnets']
        order = sorted(subs, key=lambda s: not inv.place[(v['name'].lower(), s.lower())])

        def masonry(ncol):
            cols, pos = [70] * ncol, []
            for s in order:
                items = inv.place[(v['name'].lower(), s.lower())]
                c = cols.index(min(cols))
                h = 34 if not items else 28 + ((len(items) + PER_ROW - 1) // PER_ROW) * RH
                key = (v['name'].lower(), s.lower())
                h += 22 * ((key in inv.subnet_nat) + (key in inv.subnet_rt))   # rows for the NAT / route-table markers
                pos.append((s, items, 16 + c * (SW + GAP), cols[c], h))
                cols[c] += h + 10
            return cols, pos

        # landscape layout: add subnet columns until the VNet is no taller than TARGET_VNET_H
        ncol = 1 if len(subs) <= 2 else 2
        cols, pos = masonry(ncol)
        while max(cols) > TARGET_VNET_H and ncol < len(subs):
            ncol += 1
            cols, pos = masonry(ncol)
        w = 32 + ncol * SW + (ncol - 1) * GAP
        hgt = max(cols) + 4 if subs else 80
        dns = inv.vnet_dns.get(v['name'].lower(), [])
        if dns:
            hgt += 28                                    # private DNS zones linked to this VNet
        if measure:
            return w, hgt
        name = v['name']
        bad = [_badge(t, c) for n, t, c in self.badges if n in v['peers']]
        others = [p for p in v['peers'] if p not in {b[0] for b in self.badges}]
        if 0 < len(others) <= 2:
            bad += [_badge(f'↔ {p}', '#9E9E9E') for p in others]    # few peers: name them
        elif others:
            bad.append(_badge(f'+{len(others)} peerings', '#9E9E9E'))
        if name in self.hub.values():
            style = VNET_HUB
        elif name in self.firewall_vnets:
            style = VNET_FW
        elif not v['peers']:
            style = VNET_ISOLATED
        else:
            style = VNET_SPOKE
        label = (f"<b>{name}</b><br><font style='font-size:10px'>{', '.join(v['space'])}"
                 f"{' · DNS ' + ', '.join(v['dns']) if v['dns'] else ''}</font><br>{' '.join(bad)}")
        tip = (f"Subscription: {inv.sub_name.get(v['sub'], v['sub'])}\nRegion: {v['loc']}\nAddress space: {', '.join(v['space'])}\n"
               f"DNS servers: {', '.join(v['dns']) or 'Azure default'}\nPeerings ({len(v['peers'])}):\n  " + '\n  '.join(v['peers']))
        vid = d.cell(parent, label, BOX + style + 'fontSize=12;', x, y, w, hgt, tooltip=tip)
        self.cell_of[name] = vid
        d.cell(vid, '', ICON.format(_icon('networking/Virtual_Networks')), w - 46, 8, 36, 22)
        if dns:
            zones = sorted(dns)
            shown = ', '.join(z + (' (auto-reg)' if reg else '') for z, reg in zones[:3])
            d.cell(vid, f"Private DNS zones linked ({len(zones)}): {shown}{' +' + str(len(zones) - 3) if len(zones) > 3 else ''}",
                   MARKER_RIGHT_LABEL.format(_icon('networking/DNS_Zones')), 16, hgt - 24, 18, 18,
                   tooltip='Private DNS zones linked to this VNet:\n  ' +
                           '\n  '.join(z + (' - auto-registration' if reg else '') for z, reg in zones))
        for s, items, sx, sy, h in pos:
            sid = d.cell(vid, f"<b>{s}</b> {v['prefix'].get(s, '')}", SUBNET_ON if items else SUBNET_OFF, sx, sy, SW, h)
            nsg = inv.subnet_nsg.get((name.lower(), s.lower()))
            if nsg:
                # NSG in the subnet's top-right corner, its name to the left of the icon
                d.cell(sid, nsg if len(nsg) <= 26 else nsg[:25] + '…',
                       MARKER_LEFT_LABEL.format(_icon('networking/Network_Security_Groups')), SW - 24, 4, 16, 18,
                       tooltip=f'NSG {nsg} attached to subnet {s}')
            # bottom rows: where this subnet's traffic goes next (route table) and how it reaches the internet (NAT)
            row_y = h - 22 * ((name.lower(), s.lower()) in inv.subnet_nat) - 22 * ((name.lower(), s.lower()) in inv.subnet_rt)
            rt = inv.subnet_rt.get((name.lower(), s.lower()))
            if rt:
                short, tip = inv.route_text(rt)
                d.cell(sid, short, MARKER_RIGHT_LABEL.format(_icon('networking/Route_Tables')), 8, row_y + 2, 16, 16, tooltip=tip)
                row_y += 22
            nat = inv.subnet_nat.get((name.lower(), s.lower()))
            if nat:
                d.cell(sid, f"outbound internet via NAT {nat[0]}{' (' + ', '.join(nat[1]) + ')' if nat[1] else ''}",
                       MARKER_RIGHT_LABEL.format(_icon('networking/NAT')), 8, row_y + 2, 16, 16,
                       tooltip=f"Outbound internet traffic from {s} is SNAT'ed by NAT gateway {nat[0]}"
                               f"{' to ' + ', '.join(nat[1]) if nat[1] else ''}"
                               + (" - unless a route table sends 0.0.0.0/0 elsewhere first" if rt else ''))
            for i, (lab, ic) in enumerate(items):
                r, c = divmod(i, PER_ROW)
                tip = inv.tooltips.get((name.lower(), s.lower(), lab))
                ix, iy = 8 + c * CW + (CW - 40) // 2, 26 + r * RH
                iid = d.cell(sid, lab, ICON.format(ic), ix, iy, 40, 40, tooltip=tip)
                self.cell_of[('icon', lab.split('<br>')[0].lower())] = iid
                # public IP / NIC-level NSG markers: small icons stacked beside the resource icon
                res = inv.item_res.get((name.lower(), s.lower(), lab))
                for j, (mic, mtip) in enumerate(inv.markers.get(res, [])[:4] if res else []):
                    mx, my = (ix + 42, iy + j * 20) if j < 2 else (ix - 18, iy + (j - 2) * 20)
                    d.cell(sid, '', MARKER.format(mic), mx, my, 16, 16, tooltip=mtip)
        return vid

    # ----- resource strip
    def _strip(self, parent, sid, x, y, w, measure=False):
        entries = []
        arm_order = list(ARM_STRIP)
        for arm, names in sorted(self.inv.strip[sid].items(),
                                 key=lambda kv: arm_order.index(kv[0]) if kv[0] in ARM_STRIP else 999):
            sheet = self.inv.arm_sheet.get(arm, arm)
            lab, ic = ARM_STRIP.get(arm) or STRIP.get(sheet) or (_arm_label(arm) if '/' in arm else sheet, 'general/All_Resources')
            entries.append((f'{lab} ({len(names)})', ic, f'{arm}\n' + '\n'.join(sorted(names))))
        if not entries:
            return 0
        per = max(1, (w - 16) // 110)
        h = 34 + ((len(entries) + per - 1) // per) * 78
        if measure:
            return h
        pid = self.d.cell(parent, 'Resources not attached to a subnet (private endpoints and VNet integrations are drawn inside their subnets; hover for names)',
                          BOX + 'fillColor=#FFFFFF;strokeColor=#9E9E9E;dashed=1;fontStyle=1;fontSize=11;align=center;', x, y, w, h)
        for i, (lab, ic, names) in enumerate(entries):
            r, c = divmod(i, per)
            self.d.cell(pid, lab, ICON.format(_icon(ic)), 16 + c * 110 + 35, 30 + r * 78, 40, 40, tooltip=names)
        return pid

    # ----- card panels (Log Analytics consumers, identity assignments, stale resources)
    CARD_W, CARD_H = 330, 58

    def _cards(self, parent, title, style, cards, x, y, w, measure=False):
        """cards: [(icon, name, detail, tooltip)] laid out in a grid; returns height (measure) or the panel id."""
        if not cards:
            return 0
        per = max(1, (w - 16) // (self.CARD_W + 10))
        h = 34 + ((len(cards) + per - 1) // per) * (self.CARD_H + 8)
        if measure:
            return h
        pid = self.d.cell(parent, title, BOX + style + 'fontStyle=1;fontSize=11;', x, y, w, h)
        for i, (ic, name, detail, tip) in enumerate(cards):
            r, c = divmod(i, per)
            cid = self.d.cell(pid, f"<b>{name}</b><br><font style='font-size:9px'>{detail}</font>",
                              BOX + 'fillColor=#FFFFFF;strokeColor=#BDBDBD;fontSize=10;spacingLeft=40;verticalAlign=middle;',
                              10 + c * (self.CARD_W + 10), 30 + r * (self.CARD_H + 8), self.CARD_W, self.CARD_H, tooltip=tip)
            self.d.cell(cid, '', MARKER.format(_icon(ic)), 8, self.CARD_H / 2 - 13, 26, 26)
        return pid

    MAX_CARDS = 12

    def _governance_cards(self, sid):
        inv = self.inv
        g = inv.gov.get(sid)
        cards = []
        if g and g['defender']:
            on = [n for n, t in g['defender'] if t.lower() == 'standard']
            off = [n for n, t in g['defender'] if t.lower() != 'standard']
            cards.append(('security/Azure_Defender', f"Defender for Cloud: {len(on)}/{len(g['defender'])} plans on",
                          f"off: {self._listing(off, 3)}" if off else 'all plans on',
                          'Defender plans\n  On: ' + (', '.join(on) or 'none') + '\n  Off: ' + (', '.join(off) or 'none')))
        if g and g['policy']:
            nc = sorted(g['policy'], key=lambda x: -x['noncompliant'])
            total = sum(x['noncompliant'] for x in nc)
            cards.append(('management_governance/Policy', f"Azure Policy: {len(nc)} assignment(s)",
                          f"{total} non-compliant resource evaluations" + (f" · worst: {nc[0]['name']}" if total else ''),
                          'Policy assignments (non-compliant / evaluated):\n  ' +
                          '\n  '.join(f"{x['name']}: {x['noncompliant']}/{x['evaluated']}"
                                       f"{' (not enforced)' if x['mode'].lower() == 'donotenforce' else ''}" for x in nc[:40])))
        if g and g['locks']:
            levels = collections.Counter(x['level'] for x in g['locks'])
            cards.append(('management_governance/Customer_Lockbox_for_MS_Azure', f"Resource locks: {len(g['locks'])}",
                          ', '.join(f'{v} {k}' for k, v in levels.items()),
                          'Locks:\n  ' + '\n  '.join(f"{x['level']} on {x['ttype']} {x['target']}" for x in g['locks'][:60])))
        vms = [r['Name'] for r in inv.b.sheet('Virtual Machines') if inv.sub_of_row(r) == sid
               and not str(r['Name']).lower().startswith('aks-')]
        if inv.has_backup_data:
            items = g['backup'] if g else []
            failed = [x for x in items if x['status'] and x['status'].lower() not in ('completed', 'success', 'healthy')]
            missing = [v for v in vms if v.lower() not in inv.backed_up]
            if items or missing:
                cards.append(('other/Backup_Vault', f"Backup: {len(items)} protected item(s)",
                              (f"{len(failed)} last backup not OK · " if failed else '') +
                              (f"{len(missing)} VM(s) not backed up" if missing else 'all VMs backed up'),
                              'Protected:\n  ' + '\n  '.join(f"{x['name']} ({x['vault']}): {x['status'] or x['state']}" for x in items[:40]) +
                              ('\nVMs with no backup:\n  ' + '\n  '.join(missing) if missing else '')))
        return cards

    def _cap(self, cards, what):
        """Keep panels short: the first MAX_CARDS - 1 cards, then one '+N more' card listing the rest."""
        if len(cards) <= self.MAX_CARDS:
            return cards
        keep, rest = cards[:self.MAX_CARDS - 1], cards[self.MAX_CARDS - 1:]
        return keep + [(rest[0][0], f'+{len(rest)} more {what}', 'hover for the list',
                        '\n'.join(f'{c[1]}: {c[2]}' for c in rest))]

    @staticmethod
    def _listing(items, n=4):
        items = list(items)
        shown = ', '.join(i.split(' (')[0] for i in items[:n])
        return shown + (f' +{len(items) - n} more' if len(items) > n else '')

    def _panels(self, sid):
        """(title, style, cards) for the subscription's Log Analytics, identity and stale panels."""
        inv = self.inv
        law = []          # workspaces are drawn by _monitoring (DCR -> workspace arrows)
        mi = [('identity/Managed_Identities', n, f'assigned to {len(l)}: {self._listing(l)}', 'Assigned to:\n  ' + '\n  '.join(l))
              for n, l in inv.mi_cards.get(sid, [])]
        mi = self._cap(mi, 'identities')
        # stale items grouped by kind: one card per kind, every name in the tooltip
        groups = collections.OrderedDict()
        for ic, n, why in inv.stale.get(sid, []):
            kind = STALE_KINDS.get(ic, why)
            if 'with no VM' in why:
                kind = 'Public IPs on a NIC with no VM'
            groups.setdefault(kind, (ic, []))[1].append((n, why))
        stale = [(ic, f'{kind} ({len(items)})', self._listing([n for n, _ in items], 3),
                  f'{kind}:\n  ' + '\n  '.join(f'{n} - {w}' for n, w in items)) for kind, (ic, items) in groups.items()]
        gov = self._governance_cards(sid)
        return [p for p in (
            ('Log Analytics workspaces - what sends data to each (hover for the full list)',
             'fillColor=#F4F9FF;strokeColor=#2E86C1;dashed=1;align=left;', law),
            ('User-assigned managed identities - what each is assigned to (hover for the full list)',
             'fillColor=#F7FBF4;strokeColor=#5D9C3B;dashed=1;align=left;', mi),
            ('Governance - Defender for Cloud, Azure Policy, resource locks, backup (hover for details)',
             'fillColor=#F5F3FA;strokeColor=#7B4FB5;dashed=1;align=left;', gov),
            ('Stale - not attached to anything (unattached NSGs, public IPs, NICs, disks, unassigned identities)',
             'fillColor=#FFF8F0;strokeColor=#C55A11;dashed=1;align=left;fontColor=#C55A11;', stale),
        ) if p[2]]

    # ----- monitoring: data collection rules (left) -> Log Analytics workspace (right), one arrow per DCR
    MON_ROWS = 16
    AI_SINGLE = 2      # a workspace with more Application Insights than this gets one grouped row

    def _monitoring(self, parent, sid, x, y, w, measure=False):
        """One row per data collection rule:  associated resources -> DCE (when the rule uses one) -> DCR ->
        Log Analytics workspace (one card per workspace, spanning its rules). Workspaces with no rule get a row of
        their own listing their other sources (App Insights, AKS, flow logs...)."""
        inv = self.inv
        # each workspace's rows: its data collection rules, then the Application Insights components sending to it
        def ai_rows(comps):
            # up to AI_SINGLE components get a row each; more collapse into one row so every component is counted
            if len(comps) <= self.AI_SINGLE:
                return [dict(a_, kind='ai') for a_ in comps]
            return [dict(kind='aig', name=f'{len(comps)} Application Insights', comps=comps)]
        groups = [(ws, links, [dict(d_, kind='dcr') for d_ in inv.law_dcrs.get(str(ws).lower(), [])] +
                   ai_rows(inv.law_ai.get(str(ws).lower(), [])))
                  for ws, links in inv.law_cards.get(sid, [])]
        # workspaces with App Insights first, busiest first, so the cap drops the empty ones
        groups.sort(key=lambda g_: -sum(len(x_.get('comps', [x_])) for x_ in g_[2] if x_['kind'] != 'dcr'))
        if inv.ai_classic.get(sid):
            groups.append((None, [], ai_rows(inv.ai_classic[sid])))
        if not groups:
            return 0
        def groups_dcr(ws_):
            return [x_ for g_ in groups if g_[0] == ws_ for x_ in g_[2]]
        rows, n = [], 0
        for ws, links, dcrs in groups:
            if n >= self.MON_ROWS:
                break
            ais = [x_ for x_ in dcrs if x_['kind'] != 'dcr']
            if len(dcrs) > self.MON_ROWS - n and len(ais) > 1:
                # about to be cut: fold this workspace's App Insights into one row so none are dropped
                comps = [c_ for x_ in ais for c_ in x_.get('comps', [x_])]
                dcrs = [x_ for x_ in dcrs if x_['kind'] == 'dcr'] + [
                    dict(kind='aig', name=f'{len(comps)} Application Insights', comps=comps)]
                agg = dcrs[-1:]
                dcrs = dcrs[:max(0, self.MON_ROWS - n - 1)] + agg if len(dcrs) > self.MON_ROWS - n else dcrs
                more_dcr = sum(1 for x_ in groups_dcr(ws) if x_['kind'] == 'dcr') - sum(1 for x_ in dcrs if x_['kind'] == 'dcr')
                rows.append((ws, links, dcrs, more_dcr))
                n += len(dcrs)
                continue
            take = dcrs[:self.MON_ROWS - n]
            rows.append((ws, links, take, len(dcrs) - len(take)))
            n += max(1, len(take))
        ROW = 52
        h = 34 + n * ROW + (16 if len(groups) > len(rows) else 0) + 6
        if measure:
            return h
        d = self.d
        GAPX = 40
        cw = int(max(150, min(300, (w - 20 - 3 * GAPX) / 4)))
        cols = [10 + i * (cw + GAPX) for i in range(4)]           # resources | DCE | DCR | workspace
        pid = d.cell(parent, 'Monitoring - resources -> DCE -> data collection rule / Application Insights -> Log Analytics workspace',
                     BOX + 'fillColor=#F4F9FF;strokeColor=#2E86C1;dashed=1;fontStyle=1;fontSize=11;align=left;', x, y, w, h)
        CARD = BOX + 'fillColor=#FFFFFF;strokeColor=#BDBDBD;fontSize=9;spacingLeft=34;verticalAlign=middle;'
        GREY = 'text;html=1;fontSize=8;fontColor=#9E9E9E;align=center;'
        yy = 34
        for ws, links, dcrs, more in rows:
            span = max(1, len(dcrs))
            wh = span * ROW - 10
            n_dcr = sum(1 for x_ in dcrs if x_['kind'] == 'dcr')
            n_ai = sum(len(x_.get('comps', [x_])) for x_ in dcrs if x_['kind'] != 'dcr')
            ai_names = [c_['name'] for x_ in dcrs if x_['kind'] != 'dcr' for c_ in x_.get('comps', [x_])]
            other = f'also from {len(links)}: {self._listing(links, 3)}' if links else ('' if dcrs else 'no connected resource found')
            if ws is None:
                wc = d.cell(pid, "<b>no Log Analytics workspace</b><br><font color='#B9770E'>classic Application Insights "
                                 "(retires - migrate to workspace-based)" + (f' · +{more} more' if more else '') + "</font>", CARD, cols[3], yy, cw, wh,
                            tooltip='Application Insights components that do not send to a Log Analytics workspace')
            else:
                wtip = (f'Workspace {ws}\nData collection rules: ' + (', '.join(x_['name'] for x_ in dcrs if x_['kind'] == 'dcr') or 'none') +
                        '\nApplication Insights: ' + (', '.join(ai_names) or 'none') +
                        ('\nOther sources:\n  ' + '\n  '.join(links) if links else ''))
                wc = d.cell(pid, f"<b>{ws}</b><br><font color='#555555'>{n_dcr} DCR(s){' · ' + str(n_ai) + ' App Insights' if n_ai else ''}"
                                 f"{' · ' + other if other else ''}{' · +' + str(more) + ' more' if more else ''}</font>",
                            CARD, cols[3], yy, cw, wh, tooltip=wtip)
            d.cell(wc, '', MARKER.format(_icon('analytics/Log_Analytics_Workspaces')), 6, wh / 2 - 10, 20, 20)
            for i, dcr in enumerate(dcrs):
                top = yy + i * ROW
                if dcr['kind'] == 'ai':
                    self._ai_row(pid, dcr, wc, ws, top, yy, wh, cols, cw, CARD, GREY)
                    continue
                if dcr['kind'] == 'aig':
                    self._ai_group_row(pid, dcr['comps'], wc, ws, top, yy, wh, cols, cw, CARD, GREY)
                    continue
                # data collection rule
                dtip = f"DCR {dcr['name']} -> {ws}\nCollects: {', '.join(dcr['collects']) or 'n/a'}" + (
                    f"\nEndpoint: {dcr['dce']}" if dcr['dce'] else '\nNo data collection endpoint (agents send directly)')
                dc = d.cell(pid, f"<b>{dcr['name']}</b><br><font color='#555555'>{', '.join(dcr['collects']) or 'DCR'}</font>",
                            CARD, cols[2], top, cw, 42, tooltip=dtip)
                d.cell(dc, '', MARKER.format(_icon('other/Data_Collection_Rules')), 6, 11, 20, 20)
                d.edge(dc, wc, 'sends to', 'strokeColor=#2E86C1;fontSize=8;', exit=(1, 0.5), entry=(0, (top + 21 - yy) / wh))
                # data collection endpoint
                ec = None
                if dcr['dce']:
                    ec = d.cell(pid, f"<b>{dcr['dce']}</b><br><font color='#555555'>data collection endpoint</font>",
                                CARD, cols[1], top, cw, 42, tooltip=f"DCE {dcr['dce']} - ingestion endpoint used by {dcr['name']}")
                    d.cell(ec, '', MARKER.format(_icon('management_governance/Monitor')), 6, 11, 20, 20)   # no DCE icon in draw.io
                    d.edge(ec, dc, '', 'strokeColor=#2E86C1;', exit=(1, 0.5), entry=(0, 0.5))
                # associated resources
                res = dcr['resources']
                if res:
                    names = [r_.split(' (')[0] for r_ in res]
                    rc = d.cell(pid, f"<b>{len(res)} resource(s)</b><br><font color='#555555'>{self._listing(names, 3)}</font>",
                                CARD, cols[0], top, cw, 42, tooltip='Associated with ' + dcr['name'] + ':\n  ' + '\n  '.join(res))
                    d.cell(rc, '', MARKER.format(_icon('compute/Virtual_Machine')), 6, 11, 20, 20)
                    d.edge(rc, ec or dc, '', 'strokeColor=#2E86C1;', exit=(1, 0.5), entry=(0, 0.5))
                else:
                    d.cell(pid, 'no associated resource' if inv.has_assoc else 'associations not in workbook',
                           GREY, cols[0], top + 14, cw, 14)
            yy += span * ROW
        if len(groups) > len(rows):
            hidden = groups[len(rows):]
            n_hidden_ai = sum(len(x_.get('comps', [x_])) for g_ in hidden for x_ in g_[2] if x_['kind'] != 'dcr')
            d.cell(pid, f'+{len(hidden)} more workspaces' + (f' ({n_hidden_ai} more Application Insights)' if n_hidden_ai else ''),
                   'text;html=1;fontSize=9;align=left;', 10, yy - 4, 400, 14)
        return pid

    def _ai_group_row(self, pid, comps, wc, ws, top, yy, wh, cols, cw, CARD, GREY):
        """apps (left) -> N Application Insights components (DCR column) -> workspace, for workspaces with many."""
        d = self.d
        names = sorted(c_['name'] for c_ in comps)
        apps = sorted({a_ for c_ in comps for a_ in c_['apps']})
        linked = sum(1 for c_ in comps if c_['apps'])
        tip = (f"{len(comps)} Application Insights -> {ws or 'no workspace (classic)'}\n" +
               '\n'.join(f"  {c_['name']}" + (f" <- {', '.join(c_['apps'])}" if c_['apps'] else '')
                         for c_ in sorted(comps, key=lambda c_: c_['name'])))
        ac = d.cell(pid, f"<b>{len(comps)} Application Insights</b><br><font color='#555555'>{self._listing(names, 2)}</font>",
                    CARD, cols[2], top, cw, 42, tooltip=tip)
        d.cell(ac, '', MARKER.format(_icon('devops/Application_Insights')), 6, 11, 20, 20)
        d.edge(ac, wc, 'sends to' if ws else '', 'strokeColor=#2E86C1;fontSize=8;', exit=(1, 0.5), entry=(0, (top + 21 - yy) / wh))
        if apps:
            rc = d.cell(pid, f"<b>{len(apps)} app(s)</b><br><font color='#555555'>{linked} of {len(comps)} components linked · "
                             f"{self._listing(apps, 2)}</font>", CARD, cols[0], top, cw, 42,
                        tooltip='Monitored apps:\n  ' + '\n  '.join(apps))
            d.cell(rc, '', MARKER.format(_icon('containers/App_Services')), 6, 11, 20, 20)
            d.edge(rc, ac, '', 'strokeColor=#2E86C1;', exit=(1, 0.5), entry=(0, 0.5))
        else:
            d.cell(pid, 'no app linked in the inventory', GREY, cols[0], top + 14, cw, 14)

    def _ai_row(self, pid, ai, wc, ws, top, yy, wh, cols, cw, CARD, GREY):
        """apps (left) -> Application Insights component (DCR column) -> workspace."""
        d = self.d
        kind = ai['kind'] if ai['kind'].lower() not in ('ai', 'web', '') else ''   # 'ai'/'web' are just the resource kind
        detail = ' · '.join(x for x in (kind, f"{ai['retention']} d retention" if ai['retention'] else '',
                                         f"{ai['sampling']}% sampling" if ai['sampling'] not in ('', '100', '100.0') else '') if x)
        ac = d.cell(pid, f"<b>{ai['name']}</b><br><font color='#555555'>Application Insights · {detail}</font>", CARD, cols[2], top, cw, 42,
                    tooltip=f"Application Insights {ai['name']} -> {ws or 'no workspace (classic)'}\nMonitors: " +
                            (', '.join(ai['apps']) or 'no app linked in the inventory'))
        d.cell(ac, '', MARKER.format(_icon('devops/Application_Insights')), 6, 11, 20, 20)
        d.edge(ac, wc, 'sends to' if ws else '', 'strokeColor=#2E86C1;fontSize=8;', exit=(1, 0.5), entry=(0, (top + 21 - yy) / wh))
        if ai['apps']:
            rc = d.cell(pid, f"<b>{len(ai['apps'])} app(s)</b><br><font color='#555555'>{self._listing(ai['apps'], 3)}</font>",
                        CARD, cols[0], top, cw, 42, tooltip='Monitored apps:\n  ' + '\n  '.join(ai['apps']))
            d.cell(rc, '', MARKER.format(_icon('containers/App_Services')), 6, 11, 20, 20)
            d.edge(rc, ac, '', 'strokeColor=#2E86C1;', exit=(1, 0.5), entry=(0, 0.5))
        else:
            d.cell(pid, 'no app linked in the inventory', GREY, cols[0], top + 14, cw, 14)

    # ----- subscription box (key = (subscription id, region))
    def _subscription(self, parent, key, x, y, measure=False):
        sid, loc = key
        inv = self.inv
        vs = [v for v in inv.vnets.values() if v['sub'] == sid and v['loc'] == loc]
        home = self.sub_home.get(sid) == loc
        sizes = {v['name']: self.vnet_size(v) for v in vs}
        colw = max([s[0] for s in sizes.values()] + [460])
        hubs = set(self.hub.values())
        units, narrow = [], []
        for v in vs:
            w, h = sizes[v['name']]
            if w < colw and colw >= 2 * w + GAP and v['name'] not in hubs:
                narrow.append((v, w, h))
                if len(narrow) == 2:
                    units.append(narrow)
                    narrow = []
            else:
                units.append([(v, w, h)])
        if narrow:
            units.append(narrow)
        # the hub VNet goes first, i.e. top of the left column, so its arrows leave through the subscription's
        # left edge without crossing other VNets; the rest tallest-first
        ordered = sorted(units, key=lambda u: (u[0][0]['name'] not in hubs, -max(h for _, _, h in u)))

        def masonry(k):
            cols, assigned = [56] * k, []
            for u in ordered:
                c = cols.index(min(cols))
                assigned.append((u, c, cols[c]))
                cols[c] += max(h for _, _, h in u) + 20
            return cols, assigned

        def columns(assigned, k):
            # each column is as wide as its widest unit (no dead space beside a column of narrow VNets)
            col_w = [max([sum(w for _, w, _ in u) + GAP * (len(u) - 1) for u, c, _ in assigned if c == ci] + [0])
                     for ci in range(k)]
            return col_w, sum(col_w) + COL_GAP * (k - 1)

        # landscape layout: add VNet columns while the subscription is taller than TARGET_SUB_H or taller than
        # it is wide (LANDSCAPE x width), keeping the extra column only when it actually lowers the box
        k = max(1, min(2, len(units)))
        cols, assigned = masonry(k)
        while k < len(units):
            col_w, cw = columns(assigned, k)
            if max(cols) <= TARGET_SUB_H and max(cols) <= LANDSCAPE * cw:
                break
            cols2, assigned2 = masonry(k + 1)
            if max(cols2) >= max(cols):
                break
            k, cols, assigned = k + 1, cols2, assigned2
        col_w, _ = columns(assigned, k)
        col_x = [20 + sum(col_w[:ci]) + COL_GAP * ci for ci in range(k)]
        placed = []
        for u, c, y0 in assigned:
            xo = col_x[c]
            for v, w, h in u:
                placed.append((v, xo, y0))
                xo += w + GAP
        width = max(40 + sum(col_w) + COL_GAP * (k - 1), 500)
        ph = self._strip(None, sid, 0, 0, width - 40, True) if home else 0
        panels = self._panels(sid) if home else []
        panel_h = [self._cards(None, t, st, cs, 0, 0, width - 40, True) for t, st, cs in panels]
        mon_h = self._monitoring(None, sid, 0, 0, width - 40, True) if home else 0
        height = max(cols) + ph + (mon_h + 12 if mon_h else 0) + sum(h + 12 for h in panel_h) + 20
        if not vs and not ph and not panels and not mon_h:
            height = 120
        if measure:
            return width, height
        name = inv.sub_name.get(sid, sid)
        where = '' if home else f' ({_region_title(loc)} resources)'
        nres = inv.sub_count.get(sid, 0)
        title = (f"{name}{where}<br><font style='font-size:10px;font-weight:normal'>{sid} · {len(vs)} VNet(s) here · "
                 f"{nres} resources in subscription</font>")
        bid = self.d.cell(parent, title, SUB_STYLE, x, y, width, height)
        self.cell_of[key] = bid
        for v, vx, vy in placed:
            self._vnet(bid, v, vx, vy)
        py = max(cols)
        if ph:
            self._strip(bid, sid, 20, py, width - 40)
            py += ph + 12
        if mon_h:
            self._monitoring(bid, sid, 20, py, width - 40)
            py += mon_h + 12
        for (t, st, cs), h in zip(panels, panel_h):
            self._cards(bid, t, st, cs, 20, py, width - 40)
            py += h + 12
        return bid

    def _region_boxes(self, loc):
        inv = self.inv
        subs = [s for s in inv.sub_name if any(v['sub'] == s and v['loc'] == loc for v in inv.vnets.values())
                or (self.sub_home.get(s) == loc and (inv.strip.get(s) or self._panels(s) or inv.law_cards.get(s)))
                or (loc == self.primary and not inv.sub_count.get(s))]          # assessed but empty: drawn as an empty box
        hub = self.hub.get(loc)
        hub_sub = inv.vnets[hub]['sub'] if hub else None
        subs.sort(key=lambda s: (s != hub_sub, inv.sub_name.get(s, s).lower()))
        return [(s, loc) for s in subs]

    def _region(self, title, rows, x, y, row_gap=0, bottom=20, top_band=0, padl=20):
        sizes = [[self._subscription(None, k, 0, 0, True) for k in row] for row in rows]
        w = max(padl + sum(s[0] + 30 for s in rs) for rs in sizes)
        hs = [max(s[1] for s in rs) for rs in sizes]
        h = 50 + top_band + bottom + sum(hs) + row_gap * (len(rows) - 1)
        rid = self.d.cell('1', title, REGION_STYLE, x, y, w, h)
        yo = 50 + top_band
        for row, rs, rh in zip(rows, sizes, hs):
            xo = padl
            for k, (sw, _) in zip(row, rs):
                self._subscription(rid, k, xo, yo)
                xo += sw + 30
            yo += rh + row_gap
        return rid, w, h

    # ----- whole page
    def build(self):
        inv, d = self.inv, self.d
        V = inv.vnets
        primary_hub = self.hub.get(self.primary)
        hub_key = (V[primary_hub]['sub'], self.primary) if primary_hub else None

        # primary region: hub subscription first; wrap into a second row past MAX_ROW_W
        boxes = self._region_boxes(self.primary)
        row1, width = [], 0
        for k in boxes:
            w = self._subscription(None, k, 0, 0, True)[0] + 30
            if row1 and width + w > MAX_ROW_W:
                break
            row1.append(k)
            width += w
        rows = [row1, boxes[len(row1):]] if boxes[len(row1):] else [row1]

        def peers_hub(key, hub):
            return sum(1 for v in V.values() if v['sub'] == key[0] and v['loc'] == key[1] and hub in v['peers'])

        # cross-region peerings: from the first (left-most) subscription of the primary region - the hub's, when
        # there is a hub - to every box in another region holding a VNet peered with one of its VNets. They run
        # down the region's left margin and enter the target from the top.
        src_key = row1[0]
        src_vnets = {v['name'] for v in V.values() if (v['sub'], v['loc']) == src_key}
        cross = []                                   # (region index, target key, [(source vnet, target vnet)])
        other_rows = {}
        for ri, loc in enumerate(self.regions[1:], start=1):
            other_rows[loc] = self._region_boxes(loc)
            for k in other_rows[loc]:
                pairs = [(p, v['name']) for v in V.values() if (v['sub'], v['loc']) == k for p in v['peers'] if p in src_vnets]
                if pairs:
                    cross.append((ri, k, pairs))
        n_bus = sum(1 for row in rows for k in row if k != hub_key) + 1
        padl = 20 + 8 * len(cross)

        hyb = inv.hybrid
        has_hybrid = bool(hyb['circuits'] or hyb['sites'] or hyb['conns'])
        EXT_W, LANE = (620 if inv.afd or has_hybrid else 260), 110   # wider External column holds the panels
        self.EXT_W = EXT_W
        X0 = EXT_W + 40
        TOP = 40 + LANE
        bus = LANE_STEP * (n_bus + 1) + 20 if len(rows) > 1 else 0
        n_row1 = sum(1 for k in row1 if k != hub_key and primary_hub and peers_hub(k, primary_hub))
        bottom1 = LANE_STEP * (n_row1 + 1) + 20 if len(rows) == 1 and n_row1 else 20
        rid = {}
        rid[self.primary], rw, rh = self._region(f'{_region_title(self.primary)} (primary region)', rows, X0 + 20, TOP,
                                                 row_gap=bus, bottom=bottom1, padl=padl)
        y = TOP + rh + 40
        total_w = rw
        last_region = (TOP, rw, rh)
        for ri, loc in enumerate(self.regions[1:], start=1):
            k_top = sum(1 for r, _, _ in cross if r == ri)
            row = other_rows[loc]
            sec_hub = self.hub.get(loc)
            n_sec = sum(1 for k in row[1:] if sec_hub and peers_hub(k, sec_hub))
            top_band = LANE_STEP * k_top + 16 if k_top else 0
            rid[loc], w, h = self._region(_region_title(loc), [row], X0 + 20, y, top_band=top_band,
                                          bottom=LANE_STEP * (n_sec + 1) + 20 if n_sec else 20, padl=padl)
            last_region = (y, w, h)
            y += h + 40
            total_w = max(total_w, w)
        regions_bottom = y

        # External column, top to bottom: users, Traffic Manager profiles, then one panel per Front Door profile:
        #   domains (custom domains + *.azurefd.net endpoints) -> Front Door profile -> one card per origin backend
        # Every arrow in a panel is a straight horizontal line in its own row, so none can cross.
        tm_rows = inv.b.sheet('Trafficmanagerprofiles')[:2]
        afd = inv.afd[:3]
        ROW, MAX_ROWS = 52, 8
        slots, y = [], 150
        for r in tm_rows:
            slots.append(('tm', r, y))
            y += 90
        for f in afd:
            slots.append(('afd', f, y))
            y += self._fd_height(f, ROW) + 30
        HYB_ROWS = 10
        hyb_rows = []
        if has_hybrid:
            linked = [c for c in hyb['conns'] if c['remote'] and c['gw']]
            if linked:
                hyb_rows = [('linked', c) for c in linked[:HYB_ROWS]]
            else:
                remotes = [('circuit', c) for c in hyb['circuits']] + [('site', c) for c in hyb['sites']]
                hyb_rows = [('unlinked', (remotes[i] if i < len(remotes) else None, hyb['conns'][i] if i < len(hyb['conns']) else None))
                            for i in range(min(HYB_ROWS, max(len(remotes), len(hyb['conns']))))]
            slots.append(('hyb', hyb_rows, y))
            y += 34 + max(1, len(hyb_rows)) * ROW + 40
        self.ext_h = max(330, y + 10)
        ext = d.cell('1', 'External', EXT_STYLE, 20, TOP, EXT_W - 20, self.ext_h)
        users = d.cell(ext, 'Internet users', ICON_RIGHT_LABEL.format(_icon('identity/Users')), 95, 40, 50, 50)
        tm_cells, afd_cells = [], []
        CARD = BOX + 'fillColor=#FFFFFF;strokeColor=#BDBDBD;fontSize=9;spacingLeft=34;verticalAlign=middle;'
        hyb_cards = []
        for kind, obj, sy in slots:
            if kind == 'hyb':
                hyb_cards = self._hybrid_panel(ext, obj, sy, EXT_W, ROW, CARD)
                continue
            if kind == 'tm':
                tm_cells.append((obj, d.cell(ext, f"{obj['Name']}<br>(Traffic Manager, {obj.g('Traffic Routing Method') or ''})",
                                             ICON_RIGHT_LABEL.format(_icon('networking/Traffic_Manager_Profiles')), 95, sy, 50, 50)))
                continue
            afd_cells.append(self._fd_panel(ext, obj, sy, EXT_W, ROW, CARD))

        notes = self._notes(tm_rows)
        frame = ET.Element('mxCell', id='frame', value=(
            f'<b>AZURE LANDING ZONE (current state)</b> - {len(inv.sub_name)} subscription(s) · {len(V)} VNets · '
            f'{len(self.regions)} region(s) · source: {self.source}'),
            style='rounded=1;arcSize=1;whiteSpace=wrap;html=1;verticalAlign=top;align=left;spacingLeft=10;spacingTop=4;'
                  'fillColor=none;strokeColor=#2E7D32;strokeWidth=3;fontSize=16;fontColor=#2E7D32;', parent='1', vertex='1')
        ET.SubElement(frame, 'mxGeometry', x=str(X0 - 20), y='-20', width=str(total_w + 80),
                      height=str(regions_bottom + 20), **{'as': 'geometry'})
        d.root.insert(2, frame)

        # legend + observations beside the last region when it leaves room inside the page width, else to the
        # right of everything (keeps wide estates under the ~8k px canvas limit of PNG export)
        ly, lw, _ = last_region
        if lw + 60 + 560 <= total_w:
            self._legend(X0 + 20 + lw + 60, ly, notes)
        else:
            self._legend(X0 + 20 + total_w + 60, TOP, notes)

        hubv = self.cell_of[primary_hub] if primary_hub else None
        self._external_peers(EXT_W, TOP, hubv, primary_hub)
        self._ingress(users, tm_cells, afd_cells, hyb_cards)
        if primary_hub:
            self._bus(hub_key, rows, primary_hub)
        self._cross(src_key, hubv if src_key == hub_key else None, cross, rid)
        for loc in self.regions[1:]:
            sec_hub = self.hub.get(loc)
            if sec_hub:
                self._secondary_bus(other_rows[loc], sec_hub, loc)
        self.conflicts = d.conflicts()
        return d.xml()

    FD_ROWS = 10

    @staticmethod
    def _fd_row_h(row, ROW):
        return max(1, len(row['backends']), (len(row['domains']) + 1) // 2) * ROW

    def _fd_height(self, f, ROW):
        hdr = max(1, len(f['wafs']))
        return 34 + hdr * ROW + 18 + sum(self._fd_row_h(r, ROW) for r in f['rows'][:self.FD_ROWS]) + 6

    def _fd_panel(self, ext, f, sy, EXT_W, ROW, CARD):
        """Front Door / CDN profile, one row per route:  domains (custom + endpoint) -> route -> origin backends.
        The profile and its WAF policies sit in the header row, clear of the column the users arrow comes down."""
        d = self.d
        p, pw = f['profile'], EXT_W - 40
        panel = d.cell(ext, f"{f['kind']} {p['Name']} - domains -> route -> origin backends",
                       BOX + 'fillColor=#F4F9FF;strokeColor=#2E86C1;dashed=1;fontSize=10;fontStyle=1;', 10, sy, pw, self._fd_height(f, ROW))
        DX, DW, RX, RW, BX = 10, 200, 235, 115, 380
        BW = pw - BX - 10
        hdr = max(1, len(f['wafs']))
        ph = hdr * ROW - 10
        n_routes = sum(1 for r in f['rows'] if r['route'])
        fd = d.cell(panel, f"<b>{p['Name']}</b><br><font color='#555555'>{f['kind']}<br>"
                           f"{n_routes} {'endpoint' if f['classic'] else 'route'}(s)</font>",
                    CARD.replace('spacingLeft=34', 'spacingLeft=30'), RX, 34, RW, ph,
                    tooltip=f"{f['kind']} profile {p['Name']}")
        d.cell(fd, '', MARKER.format(_icon('networking/Front_Doors')), 5, ph / 2 - 11, 22, 22)
        for i, w in enumerate(f['wafs']):
            cy = 34 + i * ROW + 21
            rules = f"{w['custom']} custom rule(s)" + (f" · {', '.join(w['managed'])}" if w['managed'] else ' · no managed rule set')
            wc = d.cell(panel, f"<b>{w['name']}</b><br><font color='{'#2E7D32' if w['mode'] == 'Prevention' else '#B9770E'}'>"
                               f"WAF · {w['mode'] or '?'}</font><font color='#555555'> · {rules}</font>",
                        CARD, BX, cy - 21, BW, 42,
                        tooltip=f"WAF policy {w['name']}: {w['mode']} mode, {w['state']}\n{rules}")
            d.cell(wc, '', MARKER.format(_icon('networking/Web_Application_Firewall_Policies_WAF')), 6, 11, 20, 20)
            d.edge(wc, fd, 'protects', 'strokeColor=#C0392B;fontSize=8;', exit=(0, 0.5), entry=(1, (cy - 34) / ph))
        y = 34 + hdr * ROW + 18
        d.cell(panel, 'routes', 'text;html=1;fontSize=9;fontColor=#555555;align=left;', RX, y - 16, 100, 14)
        first, cards = None, []
        for row in f['rows'][:self.FD_ROWS]:
            rh = self._fd_row_h(row, ROW)
            ch = rh - 10
            dc = None
            if row['domains']:
                text = '<br>'.join(f"<b>{h}</b> <font color='#555555'>{k}</font>" for h, k in row['domains'])
                dc = d.cell(panel, text, CARD.replace('verticalAlign=middle', 'verticalAlign=middle;fontSize=8'), DX, y, DW, ch,
                            tooltip='\n'.join(f'{k}: {h}' for h, k in row['domains']))
                d.cell(dc, '', MARKER.format(_icon('networking/DNS_Zones')), 6, ch / 2 - 10, 20, 20)
                first = first or dc
            rc = None
            if row['route']:
                rt = row['route']
                rc = d.cell(panel, f"<b>{rt['name']}</b><br><font color='#555555'>{rt['detail']}"
                                   f"{' -> ' + rt['group'] if rt['group'] else ''}</font>",
                            CARD.replace('spacingLeft=34', 'spacingLeft=4'), RX, y, RW, ch,
                            tooltip=f"route {rt['name']}: {rt['detail']}" + (f"\norigin group: {rt['group']}" if rt['group'] else ''))
                if dc:
                    d.edge(dc, rc, '', 'strokeColor=#1F6FC5;', exit=(1, 0.5), entry=(0, 0.5))
            else:
                d.cell(panel, 'no route in workbook', 'text;html=1;fontSize=8;fontColor=#9E9E9E;align=center;', RX, y + ch / 2 - 7, RW, 14)
            for j, o in enumerate(row['backends']):
                cy = y + j * ROW + 21
                where = f"in {o['vnet']}" if o['vnet'] else 'not in a VNet'
                bc = d.cell(panel, f"<b>{o['what']}</b><br><font color='#555555'>{o['group']} · {where}</font>",
                            CARD, BX, cy - 21, BW, 42, tooltip=f"origin {o['name']}: {o['host']}")
                d.cell(bc, '', MARKER.format(_icon(o['icon'])), 6, 11, 20, 20)
                if rc:
                    d.edge(rc, bc, '', 'strokeColor=#1F6FC5;', exit=(1, (cy - y) / ch), entry=(0, 0.5))
                cards.append((o, bc))
            if not row['backends'] and row['route']:
                d.cell(panel, 'origin group has no origin in workbook', 'text;html=1;fontSize=8;fontColor=#9E9E9E;align=left;',
                       BX, y + ch / 2 - 7, BW, 14)
            y += rh
        if not f['rows']:
            d.cell(panel, 'domains, routes and origins not in workbook - re-run the assessment',
                   'text;html=1;fontSize=9;fontColor=#9E9E9E;align=left;', DX, y, pw - 20, 14)
        # users -> first domain card (or the profile when there is none); backend cards carry any arrow into a VNet
        return dict(entry=first or fd, cards=cards)

    def _hybrid_panel(self, ext, rows, sy, EXT_W, ROW, CARD):
        """On-prem sites / ExpressRoute circuits (left) -> connection -> VNet gateway (right). Returns the right-hand
        cards with the VNet their gateway is in, so an arrow can continue into that VNet."""
        d, hyb = self.d, self.inv.hybrid
        pw = EXT_W - 40
        linked = rows and rows[0][0] == 'linked'
        title = ('Hybrid connectivity - on-prem / ExpressRoute -> connection -> VNet gateway' if linked else
                 'Hybrid connectivity - connection links not in workbook (re-run the assessment to link them)')
        panel = d.cell(ext, title, BOX + 'fillColor=#F7FBF4;strokeColor=#5D9C3B;dashed=1;fontSize=10;fontStyle=1;',
                       10, sy, pw, 34 + max(1, len(rows)) * ROW + 10)
        sites = {x['name'].lower(): ('site', x) for x in hyb['sites']}
        sites.update({x['name'].lower(): ('circuit', x) for x in hyb['circuits']})
        icon = {'site': 'networking/Local_Network_Gateways', 'circuit': 'networking/ExpressRoute_Circuits'}
        out = []
        for i, (kind, obj) in enumerate(rows):
            cy = 34 + i * ROW + 21
            if kind == 'linked':
                rk, rv = sites.get(obj['remote'].lower(), ('site', dict(name=obj['remote'], detail='')))
                left = (rk, rv)
                conn = obj
            else:
                left, conn = obj
            if left:
                lk, lv = left
                c = d.cell(panel, f"<b>{lv['name']}</b><br><font color='#555555'>{lv['detail']}</font>", CARD, 10, cy - 21, 230, 42,
                           tooltip=f"{'ExpressRoute circuit' if lk == 'circuit' else 'Local network gateway (on-prem site)'} {lv['name']}: {lv['detail']}")
                d.cell(c, '', MARKER.format(_icon(icon[lk])), 6, 11, 20, 20)
            if conn:
                ok = conn['status'].lower() in ('connected', 'succeeded', '')
                where = f" → {conn['gw']}" + (f" · {conn['vnet']}" if conn['vnet'] else '') if conn['gw'] else ''
                r = d.cell(panel, f"<b>{conn['name']}</b><br><font color='{'#555555' if ok else '#C0392B'}'>"
                                  f"{conn['type']} · {conn['status'] or 'status n/a'}{where}</font>",
                           CARD, 300, cy - 21, pw - 310, 42, tooltip=f"Connection {conn['name']} ({conn['type']}, {conn['status']})"
                                                                 + (f"\nGateway: {conn['gw']}" if conn['gw'] else '')
                                                                 + (f"\nRemote: {conn['remote']}" if conn['remote'] else ''))
                d.cell(r, '', MARKER.format(_icon('networking/Connections')), 6, 11, 20, 20)
                if kind == 'linked' and left:
                    d.edge(c, r, '', 'strokeColor=#5D9C3B;', exit=(1, 0.5), entry=(0, 0.5))
                    if conn['vnet']:
                        out.append((r, conn['vnet'], f"{conn['type']} {conn['name']}"))
        extra = len(hyb['conns']) + len(hyb['sites']) + len(hyb['circuits'])
        if extra > 2 * len(rows):
            d.cell(panel, f"{len(hyb['conns'])} connections, {len(hyb['sites'])} on-prem sites, {len(hyb['circuits'])} circuits in total",
                   'text;html=1;fontSize=9;align=left;', 10, 34 + len(rows) * ROW - 4, 400, 14)
        return out

    # ----- external peers: one column; each hub-peered one gets its own arrow to the hub VNet
    def _external_peers(self, EXT_W, TOP, hubv, hub_name):
        d, A = self.d, self.d.abs
        if not self.outside:
            return
        names = sorted(self.outside, key=lambda n: (hub_name not in self.outside[n], -len(self.outside[n]), n))
        y0 = TOP + self.ext_h + 30
        step, ch = 62, 52
        h = 44 + len(names) * step
        gid = d.cell('1', 'Peered VNets outside this inventory', BOX +
                     'fillColor=#FBE9E7;strokeColor=#C55A11;strokeWidth=2;dashed=1;fontStyle=1;fontSize=11;fontColor=#C55A11;',
                     20, y0, EXT_W - 20, h)
        linked = []
        for i, n in enumerate(names):
            peers = sorted(self.outside[n])
            c = d.cell(gid, f'<b>{n}</b><br><font style="font-size:9px">peered with {len(peers)} VNet(s)</font>',
                       BOX + 'fillColor=#FFFFFF;strokeColor=#C55A11;fontSize=10;spacingLeft=40;verticalAlign=middle;',
                       10, 36 + i * step, EXT_W - 40, ch, tooltip='Peered with:\n  ' + '\n  '.join(peers))
            d.cell(c, '', ICON.format(_icon('networking/Virtual_Networks')), 6, ch / 2 - 9, 30, 18)
            if hubv and hub_name in self.outside[n]:
                linked.append((c, A[c][1] + ch / 2, len(peers) - 1))
        if not linked:
            return
        # Each arrow: from the box's right side to its own vertical lane in the gap between the External column
        # and the region, up (or straight) to its own entry point on the hub VNet's left side. Boxes are stacked
        # top-down; upper boxes take the outer lanes and the higher entry points, so the arrows never cross.
        lo, hi = A[hubv][1] + 60, A[hubv][1] + A[hubv][3] - 10
        n = len(linked)
        estep = min(14.0, max(2.0, (hi - lo) / max(1, n)))
        lane_w = min(5.0, 30.0 / n)
        for i, (c, yc, more) in enumerate(linked):
            ye = max(lo + estep * i, min(yc, hi - estep * (n - 1 - i)))   # straight across when it fits
            xl = EXT_W + 26 + lane_w * i
            pts = [(xl, ye), (xl, yc)] if abs(ye - yc) > 0.5 else []
            d.edge(hubv, c, f'+{more} more' if more else '', 'strokeColor=#C55A11;startArrow=block;startFill=1;fontColor=#C55A11;',
                   exit=(0, d.rel_y(hubv, ye)), entry=(1, 0.5), pts=pts)

    # ----- Internet -> Traffic Manager -> the resource its endpoint IP is attached to
    def _ingress(self, users, tm_cells, afd_cells, hyb_cards=()):
        d, A, inv = self.d, self.d.abs, self.inv
        entries = [c for _, c in tm_cells] + [f['entry'] for f in afd_cells]
        if entries:
            d.edge(users, entries[0], 'DNS / HTTPS', exit=(0.5, 1), entry=(0.5, 0), label_pos=-0.6)
        ip_owner = {}
        for r in inv.b.sheet('Public IP Addresses'):
            m = re.match(r'\s*([^:]+):\s*(.+)', str(r.g('Attached To') or ''))
            if m and r.g('Ip Address'):
                ip_owner[str(r.g('Ip Address'))] = (m.group(1).strip(), m.group(2).strip())
        targets = []
        for r, tmc in tm_cells:
            for tgt in re.findall(r"'target': '([^']+)'", str(r.g('Endpoints') or '')):
                owner = ip_owner.get(tgt)
                if not owner:
                    continue
                vn = inv.res_vnet.get(('res', owner[0], owner[1].lower()))
                if vn and vn in self.cell_of:
                    targets.append((tmc, vn, f'HTTPS {tgt} -> {owner[1]}'))
        for f in afd_cells:
            for o, card in f['cards']:
                vn = o['vnet']
                if vn and vn in self.cell_of:
                    targets.append((card, vn, f"Front Door origin {o['what']}"))
        for card, vn, label in hyb_cards:
            if vn in self.cell_of:
                targets.append((card, vn, label))
        # Farther targets take the higher lane and the outer gutter so the edges nest; an edge that would still
        # touch an earlier one (sources stacked the other way round) is left out - its origins stay in the tooltip.
        targets.sort(key=lambda t: (-A[self.cell_of[t[1]]][0], A[t[0]][1]))
        drawn = 0
        for tmc, vn, label in targets:
            if drawn >= 4:
                break
            i = drawn
            vc = self.cell_of[vn]
            gut = self.EXT_W + 2 + 4 * i
            lane = 40 + 16 * i
            xg = A[vc][0] - 6 - 3 * i
            ye = A[vc][1] + 34 + 10 * i
            ty = A[tmc][1] + 25 + 5 * i
            pts = [(gut, ty), (gut, lane), (xg, lane), (xg, ye)]
            if d.would_conflict(d.anchor(tmc, (1, d.rel_y(tmc, ty))), pts, d.anchor(vc, (0, d.rel_y(vc, ye)))):
                continue
            d.edge(tmc, vc, label, exit=(1, d.rel_y(tmc, ty)), entry=(0, d.rel_y(vc, ye)), pts=pts)
            drawn += 1

    # ----- hub subscription -> every peered subscription of the primary region (band between the two rows)
    def _bus(self, hub_key, rows, hub):
        d, A, inv = self.d, self.d.abs, self.inv
        hub_sub = self.cell_of[hub_key]
        hx, hy, hw, hh = A[hub_sub]

        def n_peer(k):
            return sum(1 for v in inv.vnets.values() if v['sub'] == k[0] and v['loc'] == k[1] and hub in v['peers'])

        cx = lambda k: A[self.cell_of[k]][0] + A[self.cell_of[k]][2] / 2
        row1_bottom = max(A[self.cell_of[k]][1] + A[self.cell_of[k]][3] for k in rows[0])
        ups = sorted((k for k in rows[0] if k != hub_key and n_peer(k)), key=cx)
        downs = sorted((k for k in (rows[1] if len(rows) > 1 else []) if n_peer(k)), key=cx)
        lanes = len(ups) + len(downs)
        first_exit = hx + hw - 30 - LANE_STEP * lanes
        direct = [k for k in downs if cx(k) < first_exit - 30]
        downs = [k for k in downs if k not in direct]
        order = downs + ups[::-1]                 # exit order, left to right
        exits = {k: hx + hw - 30 - LANE_STEP * (len(order) - 1 - i) for i, k in enumerate(order)}
        depth = {k: i + 1 for i, k in enumerate(ups)}
        depth.update({k: len(ups) + 1 + i for i, k in enumerate(downs[::-1])})
        style = 'dashed=1;strokeColor=#7B4FB5;endArrow=none;'
        for k in ups + downs:
            tx, ty, tw, th = A[self.cell_of[k]]
            ly = row1_bottom + 16 + LANE_STEP * depth[k]
            d.edge(hub_sub, self.cell_of[k], f'hub peering x{n_peer(k)}', style, exit=((exits[k] - hx) / hw, 1),
                   entry=(0.5, 1 if k in ups else 0), pts=[(exits[k], ly), (tx + tw / 2, ly)])
        for k in direct:
            d.edge(hub_sub, self.cell_of[k], f'hub peering x{n_peer(k)}', style, exit=((cx(k) - hx) / hw, 1), entry=(0.5, 0))

    # ----- primary-region subscription -> boxes in other regions: down the left margin, along each region's top band
    def _cross(self, src_key, hubv, cross, rid):
        """hubv: the hub VNet when it sits in the source box - the edges then leave below it, clear of the
        external-peer arrows that use the hub VNet's own left side."""
        d, A, inv = self.d, self.d.abs, self.inv
        if not cross:
            return
        src = self.cell_of[src_key]
        cx = lambda k: A[self.cell_of[k]][0] + A[self.cell_of[k]][2] / 2
        # outermost first: later regions, then nearer targets within a region; outer edges leave higher
        ordered = sorted(cross, key=lambda e: (-e[0], cx(e[1])))
        hx, hy, hw, hh = A[src]
        y_start = A[hubv][1] + A[hubv][3] + 20 if hubv else hy + 70
        step = min(12, max(4, (hy + hh - 10 - y_start) / max(1, len(ordered))))
        primary_hub = self.hub.get(self.primary)
        by_region = collections.defaultdict(list)
        for i, (ri, k, pairs) in enumerate(ordered):
            by_region[ri].append((i, k, pairs))
        for ri, items in by_region.items():
            loc = self.regions[ri]
            ry = A[rid[loc]][1]
            # nearest target (first in items) takes the deepest lane
            for j, (i, k, pairs) in enumerate(items):
                xm = A[rid[self.primary]][0] + 4 + 8 * i
                ey = y_start + step * i
                lane = ry + 46 + LANE_STEP * (len(items) - j)
                hub_here = self.hub.get(loc)
                if hub_here and primary_hub and (primary_hub, hub_here) in pairs:
                    label = 'global VNet peering - hub to hub'
                elif len(pairs) <= 3:
                    label = 'global VNet peering: ' + ', '.join(f'{a} ↔ {b}' for a, b in pairs)
                else:
                    label = f'global VNet peering x{len(pairs)}'
                d.edge(src, self.cell_of[k], label, 'dashed=1;strokeColor=#7B4FB5;startArrow=block;startFill=1;strokeWidth=2;',
                       exit=(0, d.rel_y(src, ey)), entry=(0.5, 0), pts=[(xm, ey), (xm, lane), (cx(k), lane)],
                       label_pos=self._pos_on_last_run((A[src][0], ey), [(xm, ey), (xm, lane), (cx(k), lane)],
                                                       (cx(k), A[self.cell_of[k]][1])))

    @staticmethod
    def _pos_on_last_run(start, pts, end):
        """Label position (-1..1) at the middle of the long final horizontal run (the region's top-band lane),
        so the label sits inside the target region rather than at the path's midpoint in the margin."""
        path = [start, *pts, end]
        lengths = [abs(a[0] - b[0]) + abs(a[1] - b[1]) for a, b in zip(path, path[1:])]
        total = sum(lengths) or 1
        upto = sum(lengths[:-2]) + lengths[-2] / 2          # middle of the lane segment
        return 2 * upto / total - 1

    # ----- secondary hub subscription -> its peered subscriptions in the same region (band below the row)
    def _secondary_bus(self, row, hub, loc):
        d, A, inv = self.d, self.d.abs, self.inv
        hub_key = (inv.vnets[hub]['sub'], loc)
        if hub_key not in self.cell_of:
            return
        hub_sub = self.cell_of[hub_key]
        hx, hy, hw, hh = A[hub_sub]

        def n_peer(k):
            return sum(1 for v in inv.vnets.values() if v['sub'] == k[0] and v['loc'] == k[1] and hub in v['peers'])

        cx = lambda k: A[self.cell_of[k]][0] + A[self.cell_of[k]][2] / 2
        ups = sorted((k for k in row if k != hub_key and n_peer(k)), key=cx)
        bottom = max(A[self.cell_of[k]][1] + A[self.cell_of[k]][3] for k in row)
        for i, k in enumerate(ups):
            ex = hx + hw - 30 - LANE_STEP * i          # nearest target: rightmost exit, shallowest lane
            ly = bottom + 16 + LANE_STEP * (i + 1)
            tx, ty, tw, th = A[self.cell_of[k]]
            d.edge(hub_sub, self.cell_of[k], f'{_short(hub)} peering x{n_peer(k)}',
                   'dashed=1;strokeColor=#2E86C1;startArrow=block;startFill=1;strokeWidth=2;',
                   exit=((ex - hx) / hw, 1), entry=(0.5, 1), pts=[(ex, ly), (tx + tw / 2, ly)])

    # ----- legend + observations
    def _legend(self, lx, ly, notes):
        d = self.d
        leg = d.cell('1', 'Legend', EXT_STYLE, lx, ly, 560, 370)
        rows = [(VNET_HUB, 'Hub VNet (most-peered VNet of its region)'), (VNET_FW, 'Firewall / NVA VNet'),
                (VNET_SPOKE, 'Spoke VNet'), (VNET_ISOLATED, 'VNet with no peerings'),
                ('fillColor=#FFFFFF;strokeColor=#7FA9D9;dashed=1;', 'Subnet with resources'),
                ('fillColor=#FFFFFF;strokeColor=#C8D6E5;dashed=1;', 'Subnet with no resources in the inventory')]
        for i, (st, t) in enumerate(rows):
            d.cell(leg, '', 'rounded=1;arcSize=4;html=1;' + st, 16, 34 + i * 34, 60, 24)
            d.cell(leg, t, 'text;html=1;fontSize=11;align=left;verticalAlign=middle;', 90, 34 + i * 34, 440, 24)
        d.cell(leg, 'Peering badges: ' + ' '.join(_badge(t, c) for _, t, c in self.badges) + ' ' + _badge('+N', '#9E9E9E') +
               '<br>"(PE)" = private endpoint NIC in that subnet · hover a VNet or resource icon for details'
               "<br>NSG icon at a subnet's top-right = NSG on that subnet · small public-IP / NSG icons beside a resource ="
               ' its public IP / NIC-level NSG · "CMK: x" = VM disks encrypted with a customer-managed key (disk encryption set x)',
               'text;html=1;whiteSpace=wrap;fontSize=10;align=left;verticalAlign=top;', 16, 245, 530, 115)
        h = 60 + 34 * len(notes)
        d.cell('1', '<b>Observations from the inventory</b><ul style="margin:4px 0 0 -14px">' +
               ''.join(f'<li>{n}</li>' for n in notes) + '</ul>',
               'rounded=1;arcSize=2;whiteSpace=wrap;html=1;verticalAlign=top;align=left;spacingLeft=8;spacingTop=4;'
               'fillColor=#FFFDE7;strokeColor=#F9A825;fontSize=11;', lx, ly + 390, 560, h)

    def _notes(self, tm_rows):
        inv, V = self.inv, self.inv.vnets
        notes = []
        for loc, hub in self.hub.items():
            notes.append(f'{_region_title(loc)} hub: {hub} ({len(self.in_peers[hub])} peerings to inventoried VNets).')
        nets = []
        for n, v in V.items():
            for p in v['space']:
                try:
                    nets.append((ipaddress.ip_network(p, strict=False), n))
                except ValueError:
                    pass
        overlaps = sorted({tuple(sorted((a, b))) for (na, a), (nb, b) in
                           ((x, y) for i, x in enumerate(nets) for y in nets[i + 1:]) if a != b and na.overlaps(nb)})
        if overlaps:
            shown = ', '.join(f'{a} / {b}' for a, b in overlaps[:6])
            notes.append(f'{len(overlaps)} VNet pair(s) with overlapping address space: {shown}{" ..." if len(overlaps) > 6 else ""}.')
        lonely = [n for n, v in V.items() if not v['peers']]
        if lonely:
            notes.append(f'VNets with no peerings: {", ".join(lonely[:8])}{" ..." if len(lonely) > 8 else ""}.')
        gw = [n for n, v in V.items() if 'GatewaySubnet' in v['subnets'] and not inv.place[(n.lower(), 'gatewaysubnet')]]
        if gw:
            notes.append(f'GatewaySubnet present but no gateway in the inventory: {", ".join(gw)}.')
        dns = collections.Counter(tuple(v['dns']) for v in V.values() if v['dns'])
        if dns:
            top, cnt = dns.most_common(1)[0]
            notes.append(f'{cnt} of {len(V)} VNets use custom DNS {", ".join(top)}.')
        if inv.private_aks:
            notes.append(f'{len(inv.private_aks)} of {len(inv.b.sheet("AKS"))} AKS clusters use a private API server.')
        known = {str(r.g('Ip Address')) for r in inv.b.sheet('Public IP Addresses')}
        for r in tm_rows:
            ext = [t for t in re.findall(r"'target': '([^']+)'", str(r.g('Endpoints') or '')) if t not in known]
            if ext:
                notes.append(f"Traffic Manager {r['Name']} has endpoint(s) outside this inventory: {', '.join(ext)}.")
        no_vnet = [inv.sub_name[s] for s in inv.sub_name if not any(v['sub'] == s for v in V.values())]
        if no_vnet:
            notes.append(f'Subscriptions with no VNet: {", ".join(no_vnet)}.')
        if inv.unplaced_pe:
            notes.append('Private endpoints without a resolvable subnet: ' + ', '.join(inv.unplaced_pe[:6]) + '.')
        off = sorted({n for g in inv.gov.values() for n, t in g['defender'] if t.lower() != 'standard'})
        if off:
            notes.append(f'Defender for Cloud plans not enabled: {", ".join(off[:8])}{" ..." if len(off) > 8 else ""}.')
        if inv.has_backup_data:
            vms = [r['Name'] for r in inv.b.sheet('Virtual Machines') if not str(r['Name']).lower().startswith('aks-')]
            missing = [v for v in vms if v.lower() not in inv.backed_up]
            if missing:
                notes.append(f'{len(missing)} of {len(vms)} VMs have no backup: {", ".join(missing[:6])}{" ..." if len(missing) > 6 else ""}.')
        trunc = inv.b.sheet('Truncated Cells')
        if trunc:
            notes.append(f'{len(trunc)} workbook cell(s) were cut at Excel\'s 32,767-character limit (see the Truncated Cells sheet).')
        return notes


def build_landing_zone_drawio(file_path: str, source_name: str = None) -> str:
    """Build the complete Azure architecture draw.io XML for an assessment workbook.
    source_name is the workbook name shown in the title (defaults to the file name)."""
    inv = _Inventory(_Book(file_path))
    if not inv.vnets:
        raise ValueError('workbook has no Virtual Networks sheet rows')
    b = _Builder(inv, source_name or os.path.splitext(os.path.basename(file_path))[0])
    xml = b.build()
    if b.conflicts:
        print(f'⚠️ landing-zone diagram: {len(b.conflicts)} edge routing conflict(s): {b.conflicts[:5]}')
    return xml


if __name__ == '__main__':
    import sys
    out = build_landing_zone_drawio(sys.argv[1])
    with open(sys.argv[2], 'w', encoding='utf-8') as fh:
        fh.write(out)
    print(f'wrote {sys.argv[2]}')
