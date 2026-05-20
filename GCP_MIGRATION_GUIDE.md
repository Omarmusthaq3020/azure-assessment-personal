# GCP to Azure Migration Assessment

## Overview

The Azure Assessment Tool now supports **GCP to Azure** migration assessments in addition to Azure-to-Azure and AWS-to-Azure migrations. This feature allows you to:

- Extract all resources from a GCP project
- Map GCP resources to Azure equivalents
- Get Azure VM SKU recommendations for GCP Compute instances
- Generate comprehensive migration reports with complexity analysis

---

## Prerequisites

### 1. GCP Service Account

You need a GCP service account with the following permissions:

- **Cloud Asset Viewer** (`roles/cloudasset.viewer`)
- **Compute Viewer** (`roles/compute.viewer`)
- **Storage Object Viewer** (`roles/storage.objectViewer`)

### 2. Create Service Account

```bash
# Set your project ID
export PROJECT_ID="your-project-id"

# Create service account
gcloud iam service-accounts create azure-migration-sa \
    --display-name="Azure Migration Assessment"

# Grant permissions
gcloud projects add-iam-policy-binding $PROJECT_ID \
    --member="serviceAccount:azure-migration-sa@${PROJECT_ID}.iam.gserviceaccount.com" \
    --role="roles/cloudasset.viewer"

gcloud projects add-iam-policy-binding $PROJECT_ID \
    --member="serviceAccount:azure-migration-sa@${PROJECT_ID}.iam.gserviceaccount.com" \
    --role="roles/compute.viewer"

gcloud projects add-iam-policy-binding $PROJECT_ID \
    --member="serviceAccount:azure-migration-sa@${PROJECT_ID}.iam.gserviceaccount.com" \
    --role="roles/storage.objectViewer"

# Create and download key
gcloud iam service-accounts keys create ~/gcp-key.json \
    --iam-account=azure-migration-sa@${PROJECT_ID}.iam.gserviceaccount.com
```

---

## Usage

### Via Web UI

1. **Start the Application**
   ```bash
   cd backend
   uvicorn main:app --reload
   ```

2. **Navigate to the UI**
   - Open `http://localhost:8000`
   - Select **"GCP to Azure"** migration type
   - Click **"Proceed to Authentication"**

3. **Enter GCP Credentials**
   - **Project ID**: Your GCP project ID
   - **Service Account JSON**: Either paste the JSON or upload the key file

4. **Run Assessment**
   - Click **"Connect to GCP"**
   - The tool will extract resources and generate a report

### Via API

```python
import requests
import json

# Load service account JSON
with open('gcp-key.json', 'r') as f:
    credentials = json.load(f)

# Run assessment
response = requests.post('http://localhost:8000/api/assess-gcp', json={
    'project_id': 'your-project-id',
    'credentials_json': json.dumps(credentials)
})

result = response.json()
print(f"Report: {result['file_path']}")
```

### Via Command Line

```bash
cd backend
python gcp_assessment.py <project-id> <path-to-service-account-json>
```

---

## GCP to Azure Resource Mapping

| GCP Service | Azure Equivalent |
|-------------|------------------|
| **Compute Engine** | Virtual Machines |
| **Cloud Storage** | Blob Storage |
| **Cloud SQL** | Azure SQL Database |
| **Google Kubernetes Engine (GKE)** | Azure Kubernetes Service (AKS) |
| **BigQuery** | Azure Synapse Analytics |
| **Pub/Sub** | Azure Service Bus |
| **Cloud Functions** | Azure Functions |
| **Cloud Run** | Azure Container Apps |
| **Memorystore (Redis)** | Azure Cache for Redis |
| **Filestore** | Azure Files |
| **Bigtable / Spanner** | Azure Cosmos DB |
| **Dataproc** | Azure HDInsight |
| **Cloud Composer** | Azure Data Factory |
| **Cloud DNS** | Azure DNS |
| **VPC Network** | Virtual Network |
| **VPC Subnet** | Subnet |
| **Firewall Rules** | Network Security Group |
| **External IP** | Public IP Address |
| **Load Balancer** | Azure Load Balancer |
| **Service Account** | Managed Identity |
| **Cloud KMS** | Azure Key Vault |
| **Cloud Logging** | Azure Monitor |
| **Cloud Monitoring** | Azure Monitor Alerts |

---

## GCP Machine Type to Azure VM SKU Mapping

### E2 Series (Cost-Optimized)

| GCP Machine Type | Azure VM SKU | vCPU | RAM (GB) |
|------------------|--------------|------|----------|
| e2-micro | Standard_B1s | 2 | 1 |
| e2-small | Standard_B1ms | 2 | 2 |
| e2-medium | Standard_B2s | 2 | 4 |
| e2-standard-2 | Standard_D2s_v5 | 2 | 8 |
| e2-standard-4 | Standard_D4s_v5 | 4 | 16 |
| e2-standard-8 | Standard_D8s_v5 | 8 | 32 |

### N1 Series (General Purpose)

| GCP Machine Type | Azure VM SKU | vCPU | RAM (GB) |
|------------------|--------------|------|----------|
| n1-standard-1 | Standard_B1ms | 1 | 3.75 |
| n1-standard-2 | Standard_D2s_v5 | 2 | 7.5 |
| n1-standard-4 | Standard_D4s_v5 | 4 | 15 |
| n1-standard-8 | Standard_D8s_v5 | 8 | 30 |

### N2 Series (Balanced)

| GCP Machine Type | Azure VM SKU | vCPU | RAM (GB) |
|------------------|--------------|------|----------|
| n2-standard-2 | Standard_D2s_v5 | 2 | 8 |
| n2-standard-4 | Standard_D4s_v5 | 4 | 16 |
| n2-standard-8 | Standard_D8s_v5 | 8 | 32 |

### Compute-Optimized

| GCP Machine Type | Azure VM SKU | vCPU | RAM (GB) |
|------------------|--------------|------|----------|
| c2-standard-4 | Standard_F4s_v2 | 4 | 16 |
| c2-standard-8 | Standard_F8s_v2 | 8 | 32 |
| n2-highcpu-4 | Standard_F4s_v2 | 4 | 8 |

### Memory-Optimized

| GCP Machine Type | Azure VM SKU | vCPU | RAM (GB) |
|------------------|--------------|------|----------|
| n2-highmem-2 | Standard_E2s_v5 | 2 | 16 |
| n2-highmem-4 | Standard_E4s_v5 | 4 | 32 |
| n2-highmem-8 | Standard_E8s_v5 | 8 | 64 |

---

## Generated Report Structure

The GCP assessment generates an Excel report with the following sheets:

### 1. **Summary**
- Total resources count
- Unique services
- Unique locations
- Migration complexity breakdown

### 2. **All Resources**
- Consolidated view of all resources
- Azure service mappings
- Migration complexity ratings

### 3. **Service-Specific Sheets**
- Separate sheet for each GCP service (Compute, Storage, etc.)
- Detailed resource properties
- Flattened JSON attributes

### 4. **Migration Plan**
- Prioritized migration roadmap
- Sorted by complexity (Simple → Moderate → Complex)
- Migration notes and recommendations

---

## Migration Complexity Ratings

| Complexity | Description | Examples |
|------------|-------------|----------|
| **Simple** | Lift-and-shift migration | Compute instances, Disks, Storage buckets |
| **Moderate** | Requires configuration changes | Cloud SQL, GKE, Cloud Functions |
| **Complex** | Requires redesign or data migration | BigQuery, Spanner, Dataproc |

---

## Troubleshooting

### "Failed to access project"

**Cause**: Service account lacks permissions or project ID is incorrect.

**Solution**:
1. Verify project ID: `gcloud projects list`
2. Check service account permissions:
   ```bash
   gcloud projects get-iam-policy $PROJECT_ID \
       --flatten="bindings[].members" \
       --filter="bindings.members:serviceAccount:azure-migration-sa@*"
   ```

### "Cloud Asset Inventory API not enabled"

**Cause**: Cloud Asset API is not enabled for the project.

**Solution**:
```bash
gcloud services enable cloudasset.googleapis.com --project=$PROJECT_ID
```

### "Invalid JSON format"

**Cause**: Service account JSON is malformed or incomplete.

**Solution**:
1. Re-download the service account key
2. Ensure the entire JSON is copied (including `{` and `}`)
3. Validate JSON: `cat gcp-key.json | jq .`

---

## Limitations

1. **Single Project**: Currently supports one project per assessment
2. **No Cost Estimation**: Does not provide Azure cost estimates
3. **Manual Validation**: Complex resources require manual validation
4. **API Quotas**: Large projects may hit GCP API quotas

---

## Next Steps After Assessment

1. **Review Migration Plan**: Prioritize resources by complexity
2. **Validate Mappings**: Verify Azure SKU recommendations match workload requirements
3. **Plan Data Migration**: For databases and storage, plan data transfer strategy
4. **Test in Azure**: Create proof-of-concept in Azure for critical workloads
5. **Use Azure Migrate**: For production migrations, use Azure Migrate for orchestration

---

## Support

For issues or questions:
- Check the main README.md
- Review GCP service account permissions
- Verify Cloud Asset Inventory API is enabled
- Check backend logs for detailed error messages
