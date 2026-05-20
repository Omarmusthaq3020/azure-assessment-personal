# Azure Migration Assessment Tool

A comprehensive, AI-powered tool for assessing Azure resource migration feasibility across subscriptions and tenants. This tool analyzes resource dependencies, identifies potential blockers, and generates detailed migration plans with actionable recommendations.

---

## 📋 Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Architecture](#architecture)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Output](#output)
- [Limitations](#limitations)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## 🎯 Overview

The Azure Migration Assessment Tool is designed specifically for **Azure-to-Azure migrations** (cross-subscription or cross-tenant). It automates the complex process of:

1. **Resource Discovery** - Scans all resources in specified Azure subscriptions
2. **Dependency Mapping** - Builds a comprehensive dependency graph
3. **Move Feasibility Analysis** - Uses AI to determine if resources can be moved
4. **Migration Planning** - Generates detailed, resource-level migration plans
5. **Quota Analysis** - Provides quota usage information for target planning

### Supported Migration Scenarios

✅ **Microsoft Sponsorship → EA (Enterprise Agreement)**  
✅ **Microsoft Sponsorship → CSP (Cloud Solution Provider)**  
❌ **EA/CSP → Microsoft Sponsorship** (Not supported)  
❌ **Cross-Region Moves** (Use Azure Resource Mover instead)  
❌ **On-Premises → Azure** (Use Azure Migrate instead)

---

## ✨ Key Features

### 🔍 Intelligent Resource Analysis
- **Automated Discovery**: Queries Azure Resource Graph to fetch all resources
- **Pagination Support**: Handles environments with 1000+ resources
- **Dependency Detection**: Automatically identifies resource relationships (VMs, NICs, VNets, NSGs, etc.)
- **Generic Fallback**: Recursively scans for dependencies even for unknown resource types

### 🤖 AI-Powered Insights
- **Azure OpenAI Integration**: Analyzes each resource against Microsoft's official move documentation
- **Contextual Analysis**: Considers resource configuration, dependencies, and SKU limitations
- **Parallel Processing**: Uses multi-threading for fast analysis (up to 20 workers)
- **Retry Logic**: Handles API rate limits with exponential backoff

### 📊 Comprehensive Reporting
- **Excel Export**: Multi-sheet workbook with formatted data
- **Migration Plans**: Per-subscription migration plans with dependency tracking
- **Dashboard**: Visual summary with charts and statistics
- **Quota Information**: Region-specific quota usage for capacity planning

### 🔐 Flexible Authentication
- **Device Code Flow**: Browser-based authentication (recommended)
- **Service Principal**: Automated authentication for CI/CD
- **Isolated Sessions**: Multi-user support with session isolation

### 🌐 Modern Web Interface
- **Pre-Assessment Questionnaire**: Validates tool suitability before starting
- **Real-Time Progress**: Live progress tracking with detailed status updates
- **Responsive Design**: Works on desktop, tablet, and mobile
- **Dark Mode Support**: Modern, professional UI

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                     Frontend (React + Vite)                  │
│  - Pre-Assessment Questionnaire                             │
│  - Device Code Authentication UI                            │
│  - Progress Tracking Dashboard                              │
└──────────────────────┬──────────────────────────────────────┘
                       │ REST API
┌──────────────────────▼──────────────────────────────────────┐
│                   Backend (FastAPI)                          │
│  - Authentication Management (auth_api.py)                   │
│  - Assessment Orchestration (main.py)                        │
│  - Job Queue & Progress Tracking                            │
└──────────────────────┬──────────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────────┐
│              Core Assessment Engine (script.py)              │
│  - Azure Resource Graph Queries                             │
│  - Dependency Graph Builder                                 │
│  - Network Enrichment (VNets, Subnets, IPs)                 │
│  - Excel Report Generation                                  │
└──────────────────────┬──────────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────────┐
│         AI Analysis Engine (hybrid_analysis_functions.py)    │
│  - Dependency Fact Extraction                               │
│  - Microsoft Documentation Scraping                         │
│  - Azure OpenAI Analysis                                    │
│  - Conflict Detection                                       │
└─────────────────────────────────────────────────────────────┘
```

---

## 📦 Prerequisites

### Required Software
- **Python**: 3.10 or higher
- **Node.js**: 16.x or higher (for frontend)
- **Azure CLI**: Latest version
- **Docker** (optional): For containerized deployment

### Azure Requirements
- **Subscription Access**: Reader role minimum (Contributor recommended)
- **Azure OpenAI**: Deployed instance with GPT-4 or GPT-3.5-turbo
- **Permissions**: 
  - `Microsoft.Resources/subscriptions/read`
  - `Microsoft.ResourceGraph/resources/read`
  - `Microsoft.Network/*/read` (for network enrichment)

### Environment Variables
Create a `.env` file in the project root:

```env
# Azure OpenAI Configuration
AZURE_OPENAI_ENDPOINT=https://your-openai-instance.openai.azure.com/
AZURE_OPENAI_API_KEY=your-api-key-here
AZURE_OPENAI_DEPLOYMENT_NAME=gpt-4

# Optional: Backend Configuration
PORT=8000
```

---

## 🚀 Installation

### Option 1: Local Development Setup

#### 1. Clone the Repository
```bash
git clone <repository-url>
cd Ready-python-script
```

#### 2. Backend Setup
```bash
# Install Python dependencies
pip install -r backend/requirements.txt

# Verify Azure CLI is installed
az --version
```

#### 3. Frontend Setup
```bash
cd frontend
npm install
npm run build
cd ..
```

#### 4. Run the Application
```bash
# Start backend server
cd backend
uvicorn main:app --reload --port 8000

# In a new terminal, serve the frontend (if not using backend to serve)
cd frontend
npm run dev
```

Access the application at: `http://localhost:8000`

---

### Option 2: Docker Deployment

#### 1. Build and Run with Docker Compose
```bash
docker-compose up --build
```

#### 2. Access the Application
```
http://localhost:8000
```

#### 3. Stop the Application
```bash
docker-compose down
```

---

## ⚙️ Configuration

### Azure OpenAI Setup

1. **Deploy Azure OpenAI**:
   - Go to Azure Portal → Create Resource → Azure OpenAI
   - Deploy a model (GPT-4 recommended, GPT-3.5-turbo minimum)

2. **Get Credentials**:
   - Navigate to your Azure OpenAI resource
   - Copy the **Endpoint** and **API Key**
   - Note your **Deployment Name**

3. **Update Environment Variables**:
   ```env
   AZURE_OPENAI_ENDPOINT=https://your-instance.openai.azure.com/
   AZURE_OPENAI_API_KEY=your-key-here
   AZURE_OPENAI_DEPLOYMENT_NAME=gpt-4
   ```

### Authentication Configuration

The tool supports multiple authentication methods:

#### Device Code Flow (Recommended for Users)
- No configuration needed
- Browser-based authentication
- Supports MFA
- Session isolation per user

#### Service Principal (For Automation)
```bash
# Create a service principal
az ad sp create-for-rbac --name "migration-assessment-sp" \
  --role Reader \
  --scopes /subscriptions/{subscription-id}

# Use the output credentials in the UI or API
```

---

## 📖 Usage

### Step 1: Pre-Assessment Questionnaire

When you first access the tool, you'll answer 5 questions:

1. **Migration Type**: Select your scenario (e.g., Sponsorship → EA)
2. **Azure Access**: Confirm you have subscription access
3. **Assessment Goal**: Define what you want to achieve
4. **Reservation Limitation**: Acknowledge that reservations aren't captured
5. **Resource Count**: Estimate environment size

**Result**: The tool validates if it can meet your needs. If not, it suggests alternatives.

---

### Step 2: Authentication

#### Using Device Code (Recommended)
1. Click **"Start Device Code Authentication"**
2. Copy the provided code
3. Click the verification URL
4. Sign in with your Azure credentials
5. Enter the device code
6. Return to the tool (it auto-detects completion)

#### Using Service Principal
1. Select **"Service Principal"** tab
2. Enter:
   - Tenant ID
   - Client ID
   - Client Secret
3. Click **"Login"**

---

### Step 3: Run Assessment

1. **Select Subscriptions**: Choose which subscriptions to assess
2. **Start Assessment**: Click **"Run Assessment"**
3. **Monitor Progress**: Watch real-time progress updates:
   - Authenticating (5%)
   - Fetching resources (15-30%)
   - Building dependency graph (35-40%)
   - Analyzing with AI (45-85%)
   - Generating report (90-100%)

**Estimated Time**:
- Small (<100 resources): 2-5 minutes
- Medium (100-500): 5-15 minutes
- Large (500-2000): 15-30 minutes
- Very Large (2000+): 30+ minutes

---

### Step 4: Download & Review Report

Once complete:
1. Click **"Download Report"**
2. Open the Excel file

---

## 📄 Output

### Excel Report Structure

The generated Excel file contains multiple sheets:

#### 1. **Dashboard** (Summary)
- Total resources count
- Move feasibility breakdown (Can Move, Cannot Move, Review Required)
- Top resource types
- Subscription summary
- Visual charts

#### 2. **AllResources**
- Complete list of all discovered resources
- Columns: Name, Type, Resource Group, Location, Subscription

#### 3. **Migration Plan_{SubscriptionName}** (One per subscription)
- Resource Name
- Resource Type
- Resource Group
- Location
- **Move Status**: Can Move / Cannot Move / Review Required
- **Action Required**: Specific steps needed
- **Dependencies**: Connected resources
- **Notes**: AI-generated insights
- **Source URL**: Link to Microsoft documentation

#### 4. **{ResourceType} Sheets** (e.g., Virtual Machines, VNets)
- Detailed properties for each resource type
- Flattened JSON properties
- Network details (for VMs: Public IP, Private IP, VNet, Subnet)

#### 5. **{Subscription}_Quota**
- Region-specific quota usage
- Categories: Compute, Storage, Networking, Databases, etc.
- Current usage vs. limits

---

## ⚠️ Limitations

### What This Tool Does NOT Do

❌ **Azure Reservations**: Cannot capture reservation data  
❌ **Cost Estimation**: Does not provide migration cost estimates  
❌ **Performance Analysis**: Does not analyze performance impact  
❌ **Actual Migration**: Does not move resources (assessment only)  
❌ **Cross-Region Moves**: Not designed for region-to-region migrations  
❌ **On-Prem Migrations**: Not for on-premises to Azure scenarios  

### Known Constraints

- **API Rate Limits**: Large environments may hit Azure OpenAI rate limits (handled with retry logic)
- **Pagination**: Azure Resource Graph queries are paginated (handled automatically)
- **Dependency Coverage**: Generic fallback catches most dependencies, but complex custom setups may need manual review
- **AI Accuracy**: AI analysis is based on Microsoft documentation but should be validated for critical migrations

---

## 🐛 Troubleshooting

### Common Issues

#### 1. **"Authentication Failed"**
**Solution**:
- Verify Azure CLI is installed: `az --version`
- Check if logged in: `az account show`
- Try logging out and back in: `az logout && az login`

#### 2. **"No Resources Found"**
**Solution**:
- Verify you have Reader access to the subscription
- Check if the subscription has resources: `az resource list --subscription {id}`
- Ensure the subscription is not disabled

#### 3. **"AI Analysis Failed"**
**Solution**:
- Verify Azure OpenAI credentials in `.env`
- Check API key is valid: Test in Azure Portal
- Ensure deployment name matches your model
- Check for rate limit errors (tool retries automatically)

#### 4. **"Assessment Stuck at X%"**
**Solution**:
- Check browser console for errors (F12)
- Verify backend is running: `http://localhost:8000/health`
- Check backend logs for errors
- Large environments may take 30+ minutes

#### 5. **"Excel File is Corrupted"**
**Solution**:
- Ensure `openpyxl` is installed: `pip install openpyxl`
- Check disk space
- Try re-running the assessment

---

## 🔧 Advanced Configuration

### Parallel Processing Tuning

Edit `backend/hybrid_analysis_functions.py`:

```python
# Line 756-763: Adjust worker count based on environment size
if total_resources < 20:
    max_workers = 5  # Conservative for small
elif total_resources < 100:
    max_workers = 10  # Medium
elif total_resources < 300:
    max_workers = 15  # Large
else:
    max_workers = 20  # Very large
```

### Custom Resource Type Handling

Add custom dependency logic in `backend/hybrid_analysis_functions.py`:

```python
# Line 100+: Add your custom resource type
elif res_type == "microsoft.yourservice/yourtype":
    # Extract dependencies
    custom_dependency = res.get("properties", {}).get("linkedResource")
    if custom_dependency:
        dependencies.append(custom_dependency.lower())
```

---

## 📊 Performance Optimization

### For Large Environments (1000+ resources)

1. **Increase Worker Count**: Edit `max_workers` in `hybrid_analysis_functions.py`
2. **Use Faster AI Model**: Switch to GPT-3.5-turbo for speed (less accurate)
3. **Disable AI Analysis**: Set `requiresAI: false` in questionnaire
4. **Run During Off-Peak**: Avoid Azure API throttling

### For Slow Networks

1. **Increase Timeout**: Edit `timeout=30` in `move_support_fetcher.py`
2. **Use Docker**: Reduces network hops
3. **Run on Azure VM**: Closer to Azure APIs

---

## 🤝 Contributing

We welcome contributions! Please follow these guidelines:

1. **Fork the Repository**
2. **Create a Feature Branch**: `git checkout -b feature/your-feature`
3. **Commit Changes**: `git commit -m "Add your feature"`
4. **Push to Branch**: `git push origin feature/your-feature`
5. **Open a Pull Request**

### Code Style
- **Python**: Follow PEP 8
- **JavaScript**: Use ESLint configuration
- **Comments**: Document complex logic

---

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

---

## 🙏 Acknowledgments

- **Microsoft Azure**: For Azure Resource Graph and Azure OpenAI
- **OpenAI**: For GPT models
- **FastAPI**: For the backend framework
- **React**: For the frontend framework

---

**Last Updated**: December 2025  
**Version**: 1.0.0
