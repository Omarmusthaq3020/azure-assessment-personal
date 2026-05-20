"""
Azure Assessment Tool - Configuration File
============================================
This file contains all configurable parameters for the Azure Assessment Tool.
Modify values here instead of searching through the codebase.
"""

import os
from typing import Dict, List


# ============================================================================
# APPLICATION SETTINGS
# ============================================================================

class AppConfig:
    """Application-level configuration"""
    
    # Server Configuration
    HOST = "0.0.0.0"
    PORT = int(os.getenv("PORT", "8000"))
    
    # CORS Settings
    CORS_ALLOW_ORIGINS = [
        origin.strip()
        for origin in os.getenv("CORS_ALLOW_ORIGINS", "http://localhost:5173").split(",")
        if origin.strip()
    ]
    CORS_ALLOW_CREDENTIALS = True
    CORS_ALLOW_METHODS = ["*"]
    CORS_ALLOW_HEADERS = ["*"]
    
    # File Paths
    PROVIDERS_RAW_JSON = "providers_raw.json"
    FRONTEND_DIST_DIR = "frontend/dist"
    DATA_DIR = "data"
    
    # Temporary Directory Settings
    TEMP_DIR_PREFIX = "az_session_"


# ============================================================================
# AZURE OPENAI SETTINGS
# ============================================================================

class AzureOpenAIConfig:
    """Azure OpenAI API configuration"""
    
    # Azure OpenAI Endpoint
    ENDPOINT = os.getenv(
        "AZURE_OPENAI_ENDPOINT",
        "https://azure-migration-open-ai.openai.azure.com/"
    )
    
    # Azure OpenAI API Key — set via AZURE_OPENAI_API_KEY env var, never hardcode
    API_KEY = os.getenv("AZURE_OPENAI_API_KEY", "")
    if not API_KEY:
        print("⚠️ WARNING: AZURE_OPENAI_API_KEY is missing or empty!")
    
    # Deployment Name
    DEPLOYMENT_NAME = os.getenv(
        "AZURE_OPENAI_DEPLOYMENT_NAME",
        "gpt-4.1"
    )
    
    # API Version
    API_VERSION = "2024-10-21"
    
    # Retry Configuration
    MAX_RETRIES = 5
    INITIAL_RETRY_DELAY = 2  # seconds
    MAX_RETRY_DELAY = 60  # seconds
    BACKOFF_MULTIPLIER = 2
    
    # Token Limits
    MAX_TOKENS_PER_REQUEST = 500000
    TEMPERATURE = 0.3


# ============================================================================
# AZURE RESOURCE GRAPH SETTINGS
# ============================================================================

class AzureResourceGraphConfig:
    """Azure Resource Graph query configuration"""

    # Query Settings
    BATCH_SIZE = 1000  # Number of resources to fetch per batch

    QUERY = "Resources"

    # Column cleanup: drop noisy / duplicate columns produced by json_normalize.
    # Match is done against the *leaf* name (last dotted segment, lowercased).
    NOISE_COLUMN_LEAVES = {
        "provisioningstate", "etag", "tenantid", "vmid", "kind", "managedby",
        "createdtime", "changedtime", "uniqueid", "guid", "version",
        "ipconfigurationsetag", "subnetsetag", "securityrulesetag",
        "principalid", "clientid",  # duplicated across identity blocks
    }

    # Drop columns whose full dotted path contains any of these substrings.
    NOISE_COLUMN_PATH_SUBSTRINGS = (
        "extendedlocation", "systemdata", "additionalcapabilities",
        ".id",  # nested object id refs (we already have top-level id)
    )

    # Maximum nesting depth to flatten. Properties deeper than this are kept
    # as a single JSON string column instead of being exploded into many cols.
    MAX_FLATTEN_DEPTH = 2

    # Resource Exclusions — substring match against the lowered resource type.
    # Covers Azure Portal "hidden" types (the ones filtered out when "Show hidden
    # types" is unchecked in the Portal) plus a few internal/auxiliary types.
    EXCLUDE_KEYWORDS = [
        # Pre-existing
        "movecollection",
        "networkwatcher",

        # Monitoring / alerting (hidden in Portal)
        "smartdetectoralertrules",
        "/autoscalesettings",
        "/metricalerts",
        "/scheduledqueryrules",
        "/actiongroups",
        "/activitylogalerts",
        "/datacollectionrules",
        "/datacollectionendpoints",
        "/webtests",

        # Compute auxiliary (hidden in Portal)
        "/virtualmachines/extensions",
        "/virtualmachines/runcommands",
        "/restorepointcollections",
        "/galleries/images",
        "/sshpublickeys",

        # Network auxiliary (hidden in Portal)
        "/networkwatchers/",

        # Migration / discovery (hidden in Portal)
        "microsoft.migrate/",

        # Security / Defender (hidden in Portal)
        "microsoft.security/",

        # Support (hidden in Portal)
        "microsoft.support/",

        # Azure AD / Entra (hidden in Portal)
        "microsoft.aad/",
        "microsoft.azureactivedirectory/",

        # Recovery Services internals (hidden in Portal)
        "/backupfabrics/",

        # Log Analytics / DevTest internals (hidden in Portal)
        "/savedsearches",
        "microsoft.devtestlab/schedules",
    ]


# ============================================================================
# AZURE QUOTA SETTINGS
# ============================================================================

class AzureQuotaConfig:
    """Azure Quota API configuration"""
    
    # Provider API Versions
    PROVIDER_API_VERSIONS: Dict[str, str] = {
        "Microsoft.Compute": "2024-07-01",
        "Microsoft.Network": "2023-11-01",
        "Microsoft.Storage": "2023-01-01",
        "Microsoft.MachineLearningServices": "2023-10-01",
    }
    
    # Unsupported Regions per Provider
    UNSUPPORTED_REGIONS: Dict[str, List[str]] = {
        "Microsoft.MachineLearningServices": ["westindia", "northindia"]
    }
    
    # Quota Categories and Keywords
    QUOTA_CATEGORIES = {
        "Compute": [
            "vcpus", "cores", "virtual machines", "dedicated", "low-priority", "spot",
            "availability set", "proximity placement", "host group", "scale set", "gpu"
        ],
        "Storage": [
            "storage", "disk", "unmanaged", "managed", "blob", "queue", "file", "table",
            "standardssd", "premium storage", "snapshots"
        ],
        "Networking": [
            "ip address", "load balancer", "public ip", "private ip", "network interface",
            "virtual network", "gateway", "vnet", "subnet", "dns", "expressroute", "vpn",
            "application gateway", "firewall", "nat gateway"
        ],
        "Databases": [
            "sql", "mysql", "postgres", "cosmos", "mariadb", "elastic pool", "database"
        ],
        "Containers": [
            "kubernetes", "aks", "container instance", "container group", "container registry"
        ],
        "AI/ML": [
            "cognitive services", "openai", "ml workspace", "bot service"
        ],
        "Integration": [
            "logic app", "function", "event grid", "event hub", "service bus", "data factory",
            "data integration"
        ],
        "Security": [
            "key vault", "security center", "defender", "sentinel", "policy", "identity",
            "access"
        ],
        "Monitoring": [
            "log analytics", "application insights", "monitor", "metrics", "alerts",
            "autoscalesetting", "action group"
        ],
        "Web & App Services": [
            "app service", "app plan", "web app", "api management", "static site", "function app"
        ]
    }
    
    # Parallel Processing
    MAX_WORKERS = 4


# ============================================================================
# EXCEL EXPORT SETTINGS
# ============================================================================

class ExcelExportConfig:
    """Excel export and formatting configuration"""
    
    # File Naming
    FILE_NAME_PREFIX = "azure_resources_detailed"
    FILE_NAME_DATE_FORMAT = "%Y%m%d_%H%M%S"
    
    # Sheet Name Settings
    MAX_SHEET_NAME_LENGTH = 31
    INVALID_SHEET_NAME_CHARS = r'[\\/*?:\[\]]'
    SHEET_NAME_REPLACEMENT = '_'
    
    # Styling
    HEADER_COLOR = "00B0F0"  # Blue
    HEADER_FONT_COLOR = "FFFFFF"  # White
    BORDER_STYLE = "thin"
    
    # Column Width Settings
    DEFAULT_COLUMN_PADDING = 2
    MAX_COLUMN_WIDTH = 50
    NOTES_COLUMN_MAX_WIDTH = 100
    ACTION_REQUIRED_COLUMN_WIDTH = 100
    SOURCE_URL_COLUMN_WIDTH = 100
    
    # Data Limits
    MAX_EMPTY_ROWS_TO_CHECK = 10
    
    # Migration Plan Columns
    MIGRATION_PLAN_HEADERS = [
        "Resource Name", "Resource Type", "Resource Group", "Location",
        "Move Status", "Action Required", "Dependencies", "Notes", "Source URL"
    ]


# ============================================================================
# RESOURCE TYPE FRIENDLY NAMES
# ============================================================================

class ResourceTypeFriendlyNames:
    """Override friendly names for specific resource types"""
    
    OVERRIDES: Dict[str, str] = {
        "microsoft.dbformysql/flexibleservers": "MySQL Flexible Servers",
        "microsoft.dbforpostgresql/flexibleservers": "PostgreSQL Flexible Servers",
        "microsoft.insights/metricalerts": "Metric Alerts",
        "microsoft.insights/scheduledqueryrules": "Scheduled Query Rules",
        "microsoft.datafactory/factories": "Data Factories",
        "microsoft.network/dnszones": "DNS Zone",
        "microsoft.insights/actiongroups": "Action Group",
        "microsoft.network/publicipaddresses": "Public IP Addresses",
        "microsoft.insights/components": "Application Insights",
        "microsoft.web/serverfarms": "App Service Plan",
        "microsoft.web/sites": "App Service",
        "microsoft.containerregistry/registries": "Azure Container Registry",
        "microsoft.containerservice/managedclusters": "AKS",
        "microsoft.saas/resources": "Saas_Resources",
        "microsoft.keyvault/vaults": "Key Vault",
        "microsoft.cognitiveservices/accounts": "Cognitive Services",
        "microsoft.recoveryservices/vaults": "Recovery Services Vault",
        
        # AWS to Azure Service Mappings (Added for multi-cloud assessments)
        "aws::ec2::snapshot": "Azure Snapshots",
        "aws::ec2::securitygroupingress": "Assess Manually",
        "aws::ec2::securitygroupegress": "Assess Manually",
        "aws::ec2::volume": "Managed Disks",
        "aws::ec2::networkinterface": "Network Interface (NIC)",
        "aws::ec2::instance": "Virtual Machines",
        "aws::ec2::keypair": "Azure SSH Keys",
        "aws::s3::bucket": "Storage Accounts",
        "aws::ec2::fleet": "Assess Manually",
        "aws::iam::role": "Microsoft Entra ID",
        "aws::ec2::spotinstancerequest": "Assess Manually",
        "aws::dlm::lifecyclepolicy": "Assess Manually",
        "aws::secretsmanager::secret": "Azure Key Vault Secrets",
        "aws::mgn::sourceserver": "Azure Migrate Source Server",
        "aws::ec2::securitygroup": "Network Security Groups",
        "aws::cloudformation::stack": "Azure Resource Manager (ARM)",
        "aws::iam::user": "Microsoft Entra ID",
        "aws::kms::key": "Azure Key Vault",
        "aws::ec2::subnet": "Virtual Network",
        "aws::ec2::routetable": "Route Tables",
        "aws::ec2::eip": "Public IP Addresses (Static)",
        "aws::ec2::internetgateway": "N/A (Implicit in VNet)",
        "aws::lightsail::staticip": "Public IP Addresses (Static)",
        "aws::ec2::vpc": "Virtual Network",
        "aws::ec2::vpcflowlog": "Assess Manually",
        "aws::ec2::launchtemplate": "VM Scale Set / Image Template",
        "aws::cloudfront::cachepolicy": "Assess Manually",
        "aws::cloudfront::distribution": "Azure Front Door / CDN",
        "aws::ec2::networkacl": "Network Security Groups",
        "aws::ec2::networkinterfaceattachment": "N/A (Implicit in NIC)",
        "aws::ec2::dhcpoptions": "VNet DNS Settings",
        "aws::cloudwatch::alarm": "Azure Monitor Alerts",
        "aws::cloudfront::responseheaderspolicy": "Assess Manually",
        "aws::cloudfront::originrequestpolicy": "Assess Manually",
        "aws::ec2::prefixlist": "IP Groups",
        "aws::iam::managedpolicy": "Azure RBAC / Policies",
        "aws::iam::instanceprofile": "Managed Identity",
        "aws::events::rule": "Azure Event Grid",
        "aws::elasticache::parametergroup": "Azure Cache for Redis",
        "aws::ecs::cluster": "Azure Kubernetes Service (AKS)",
        "aws::ec2::gatewayattachment": "Assess Manually",
        "aws::ecs::capacityprovider": "Azure Container Apps / AKS",
        "aws::ec2::vpcendpoint": "Private Endpoints / Private Link",
        "aws::ec2::vpcdhcpoptsassociation": "VNet DNS Settings",
        "aws::elasticache::user": "Azure Cache for Redis",
        "aws::kms::alias": "Azure Key Vault (Alias)",
        "aws::sns::topic": "Azure Service Bus",
        "aws::iam::virtualmfadevice": "Entra ID MFA",
        "aws::lightsail::keypair": "Azure SSH Keys",
        "aws::networkfirewall::statelessrulegroup": "Assess Manually",
    }


# ============================================================================
# AKS DETECTION SETTINGS
# ============================================================================

class AKSDetectionConfig:
    """Configuration for detecting AKS-related resources"""
    
    # Resource types that might be AKS-related
    AKS_RELATED_RESOURCE_TYPES = [
        "microsoft.network/loadbalancers",
        "microsoft.network/publicipaddresses",
        "microsoft.network/networkinterfaces"
    ]
    
    # Keywords to detect AKS resources
    AKS_KEYWORDS = ["aks", "kubernetes"]
    
    # Friendly name for AKS resources
    AKS_RESOURCE_LABEL = "AKS Resource"


# ============================================================================
# ASSESSMENT JOB PROGRESS TRACKING
# ============================================================================

class AssessmentProgressConfig:
    """Configuration for assessment job progress tracking"""
    
    # Progress Events Mapping (keyword -> (percentage, step, message))
    PROGRESS_EVENTS = {
        "Fetching subscription names": (5, "Authenticating", "Fetching subscriptions..."),
        "Querying Azure Resource Graph": (15, "Fetching resources", "Querying Azure Resource Graph..."),
        "Building NIC to Public IP map": (25, "Fetching resources", "Building network mappings..."),
        "Enriching VM resources": (30, "Fetching resources", "Enriching resource details..."),
        "Loading friendly type mappings": (35, "Building dependency graph", "Loading resource metadata..."),
        "Detecting AKS-related resources": (38, "Building dependency graph", "Detecting resource relationships..."),
        "Exporting to Excel": (40, "Building dependency graph", "Analyzing dependencies..."),
        "Formatting Excel file": (42, "Generating report", "Formatting Excel..."),
        "Adding migration plan with dependency analysis": (45, "Analyzing with AI", "Starting AI analysis..."),
        "Appending quota info": (85, "Generating report", "Adding quota information..."),
        "No access to view reservations": (88, "Generating report", "Finalizing report..."),
        "Appending quota info based on usage": (90, "Creating dashboard", "Creating dashboard..."),
    }
    
    # AI Analysis Progress
    AI_ANALYSIS_START_PERCENTAGE = 45
    AI_ANALYSIS_END_PERCENTAGE = 85
    AI_ANALYSIS_WEIGHT = 40  # Percentage points allocated to AI analysis


# ============================================================================
# NETWORK SETTINGS
# ============================================================================

class NetworkConfig:
    """Network-related configuration"""
    
    # Public IP fetch settings
    PUBLIC_IP_LIST_COMMAND = ["az", "network", "public-ip", "list", "--output", "json"]
    
    # NIC show command template
    NIC_SHOW_COMMAND_TEMPLATE = ["az", "network", "nic", "show", "--ids", "{nic_id}", "--output", "json"]
    
    # Subnet show command template
    SUBNET_SHOW_COMMAND_TEMPLATE = ["az", "network", "vnet", "subnet", "show", "--ids", "{subnet_id}", "--output", "json"]


# ============================================================================
# AI ANALYSIS SETTINGS
# ============================================================================

class AIAnalysisConfig:
    """AI analysis configuration"""
    
    # Text Limits
    MAX_CHARS_PER_SHEET = 10000
    MAX_COMBINED_SUMMARY_CHARS = 50000
    MAX_DATA_CONTEXT_CHARS = 20000
    
    # Chat History
    MAX_CHAT_HISTORY_MESSAGES = 10
    
    # System Prompts
    SYSTEM_PROMPT = "You are a helpful Azure Migration Architect."
    
    # Analysis Prompt Template
    ANALYSIS_PROMPT_TEMPLATE = """
You are an Azure Migration Expert. Analyze the following comprehensive assessment data from multiple sheets:

{combined_summary}

        Please provide a high-level executive summary in Markdown format, including:
1. **Executive Summary**: Overall readiness, total resources, and key statistics.
2. **Risk Analysis**: Identify high-risk resources and blockers.
3. **Migration Strategy**: Recommended approach and priorities.
4. **Next Steps**: Actionable items for the team.

        Keep it concise and professional.
        """


# ============================================================================
# PERFORMANCE SETTINGS
# ============================================================================

class PerformanceConfig:
    """Runtime toggles for performance-sensitive analysis stages."""

    # If true, low-risk resources (no blockers detected) skip AI and are marked movable.
    SKIP_AI_FOR_LOW_RISK = os.getenv("SKIP_AI_FOR_LOW_RISK", "true").lower() == "true"
    ENABLE_IAM_ANALYSIS = os.getenv("ENABLE_IAM_ANALYSIS", "true").lower() == "true"
    ENABLE_QUOTA_ANALYSIS = os.getenv("ENABLE_QUOTA_ANALYSIS", "true").lower() == "true"
    ENABLE_IAM_SHEET = os.getenv("ENABLE_IAM_SHEET", "true").lower() == "true"


# ============================================================================
# MOVE SUPPORT FETCHER SETTINGS
# ============================================================================

class MoveSupportFetcherConfig:
    """Configuration for move support data fetching"""
    
    # Text truncation limit
    MAX_TEXT_CHARS = 8000
    CACHE_FILE_PATH = os.getenv(
        "MOVE_SUPPORT_CACHE_FILE_PATH",
        "/app/data/move_support_cache.json" if os.path.isdir("/app/data") else "backend/move_support_cache.json"
    )
    CACHE_DURATION_DAYS = int(os.getenv("MOVE_SUPPORT_CACHE_DURATION_DAYS", "7"))
    FOLLOW_DETAIL_LINKS = os.getenv("MOVE_SUPPORT_FOLLOW_DETAIL_LINKS", "false").lower() == "true"


# ============================================================================
# NON-MOVABLE RESOURCES SETTINGS
# ============================================================================

class NonMovableResourcesConfig:
    """Configuration for non-movable resources management"""
    
    # Cache Settings
    CACHE_DURATION_DAYS = 7  # How long to cache before re-scraping
    
    # File Paths
    JSON_FILE_PATH = os.getenv(
        "NON_MOVABLE_JSON_FILE_PATH",
        "/app/data/non_movable_resources.json" if os.path.isdir("/app/data") else "backend/non_movable_resources.json"
    )
    
    # Source URL (same as move support fetcher)
    SOURCE_URL = "https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources"
    
    # Auto-update Settings
    AUTO_UPDATE_ENABLED = os.getenv("NON_MOVABLE_AUTO_UPDATE_ENABLED", "false").lower() == "true"
    
    # Scraper Settings
    REQUEST_TIMEOUT = 30  # seconds
    RETRY_ON_FAILURE = True


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_config_summary() -> str:
    """Returns a summary of current configuration"""
    return f"""
Azure Assessment Tool Configuration Summary
==========================================

Application:
  - Host: {AppConfig.HOST}
  - Port: {AppConfig.PORT}

Azure OpenAI:
  - Endpoint: {AzureOpenAIConfig.ENDPOINT}
  - Deployment: {AzureOpenAIConfig.DEPLOYMENT_NAME}
  - API Version: {AzureOpenAIConfig.API_VERSION}

Resource Graph:
  - Batch Size: {AzureResourceGraphConfig.BATCH_SIZE}
  - Excluded Keywords: {', '.join(AzureResourceGraphConfig.EXCLUDE_KEYWORDS)}

Quota Providers:
  - {len(AzureQuotaConfig.PROVIDER_API_VERSIONS)} providers configured
  - Max Workers: {AzureQuotaConfig.MAX_WORKERS}

Excel Export:
  - File Prefix: {ExcelExportConfig.FILE_NAME_PREFIX}
  - Max Sheet Name Length: {ExcelExportConfig.MAX_SHEET_NAME_LENGTH}
"""


if __name__ == "__main__":
    print(get_config_summary())

