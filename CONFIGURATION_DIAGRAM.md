# Configuration Architecture

```
Azure Assessment Tool - Configuration System
═══════════════════════════════════════════════

┌─────────────────────────────────────────────────────────────┐
│                    CONFIGURATION LAYER                       │
└─────────────────────────────────────────────────────────────┘
                              │
                              │
        ┌─────────────────────┼─────────────────────┐
        │                     │                     │
        ▼                     ▼                     ▼
┌──────────────┐      ┌──────────────┐     ┌──────────────┐
│   .env       │      │  config.py   │     │ docker-      │
│  (Secrets)   │      │ (Settings)   │     │ compose.yml  │
└──────────────┘      └──────────────┘     └──────────────┘
        │                     │                     │
        │                     │                     │
        └─────────────────────┼─────────────────────┘
                              │
                              ▼
                    ┌──────────────────┐
                    │   Application    │
                    │   Components     │
                    └──────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        │                     │                     │
        ▼                     ▼                     ▼
┌──────────────┐      ┌──────────────┐     ┌──────────────┐
│   Backend    │      │   Frontend   │     │   Scripts    │
│   (FastAPI)  │      │   (React)    │     │   (Python)   │
└──────────────┘      └──────────────┘     └──────────────┘


Configuration Flow
══════════════════

1. Environment Variables (.env)
   ↓
2. Config Classes (config.py)
   ↓
3. Application Code
   ↓
4. Runtime Behavior


Configuration Classes Hierarchy
════════════════════════════════

config.py
│
├── AppConfig
│   ├── HOST
│   ├── PORT
│   ├── CORS_ALLOW_ORIGINS
│   ├── CORS_ALLOW_CREDENTIALS
│   ├── CORS_ALLOW_METHODS
│   ├── CORS_ALLOW_HEADERS
│   ├── PROVIDERS_RAW_JSON
│   ├── FRONTEND_DIST_DIR
│   ├── DATA_DIR
│   └── TEMP_DIR_PREFIX
│
├── AzureOpenAIConfig
│   ├── ENDPOINT
│   ├── API_KEY
│   ├── DEPLOYMENT_NAME
│   ├── API_VERSION
│   ├── MAX_RETRIES
│   ├── INITIAL_RETRY_DELAY
│   ├── MAX_RETRY_DELAY
│   ├── BACKOFF_MULTIPLIER
│   ├── MAX_TOKENS_PER_REQUEST
│   └── TEMPERATURE
│
├── AzureResourceGraphConfig
│   ├── BATCH_SIZE
│   ├── QUERY
│   └── EXCLUDE_KEYWORDS
│
├── AzureQuotaConfig
│   ├── PROVIDER_API_VERSIONS
│   ├── UNSUPPORTED_REGIONS
│   ├── QUOTA_CATEGORIES
│   └── MAX_WORKERS
│
├── ExcelExportConfig
│   ├── FILE_NAME_PREFIX
│   ├── FILE_NAME_DATE_FORMAT
│   ├── MAX_SHEET_NAME_LENGTH
│   ├── INVALID_SHEET_NAME_CHARS
│   ├── SHEET_NAME_REPLACEMENT
│   ├── HEADER_COLOR
│   ├── HEADER_FONT_COLOR
│   ├── BORDER_STYLE
│   ├── DEFAULT_COLUMN_PADDING
│   ├── MAX_COLUMN_WIDTH
│   ├── NOTES_COLUMN_MAX_WIDTH
│   ├── ACTION_REQUIRED_COLUMN_WIDTH
│   ├── SOURCE_URL_COLUMN_WIDTH
│   ├── MAX_EMPTY_ROWS_TO_CHECK
│   └── MIGRATION_PLAN_HEADERS
│
├── ResourceTypeFriendlyNames
│   └── OVERRIDES
│
├── AKSDetectionConfig
│   ├── AKS_RELATED_RESOURCE_TYPES
│   ├── AKS_KEYWORDS
│   └── AKS_RESOURCE_LABEL
│
├── AssessmentProgressConfig
│   ├── PROGRESS_EVENTS
│   ├── AI_ANALYSIS_START_PERCENTAGE
│   ├── AI_ANALYSIS_END_PERCENTAGE
│   └── AI_ANALYSIS_WEIGHT
│
├── NetworkConfig
│   ├── PUBLIC_IP_LIST_COMMAND
│   ├── NIC_SHOW_COMMAND_TEMPLATE
│   └── SUBNET_SHOW_COMMAND_TEMPLATE
│
├── AIAnalysisConfig
│   ├── MAX_CHARS_PER_SHEET
│   ├── MAX_COMBINED_SUMMARY_CHARS
│   ├── MAX_DATA_CONTEXT_CHARS
│   ├── MAX_CHAT_HISTORY_MESSAGES
│   ├── SYSTEM_PROMPT
│   └── ANALYSIS_PROMPT_TEMPLATE
│
└── MoveSupportFetcherConfig
    └── MAX_TEXT_CHARS


Environment Variables Priority
═══════════════════════════════

Highest Priority
    ↓
1. Docker Compose environment section
    ↓
2. .env file
    ↓
3. System environment variables
    ↓
4. config.py defaults
    ↓
Lowest Priority


Usage Examples
══════════════

Example 1: Using in Python Code
────────────────────────────────
from config import AzureOpenAIConfig

client = AzureOpenAI(
    azure_endpoint=AzureOpenAIConfig.ENDPOINT,
    api_key=AzureOpenAIConfig.API_KEY,
    api_version=AzureOpenAIConfig.API_VERSION
)


Example 2: Using in Docker
───────────────────────────
version: '3.8'
services:
  app:
    environment:
      - PORT=${PORT}
      - AZURE_OPENAI_ENDPOINT=${AZURE_OPENAI_ENDPOINT}
      - AZURE_OPENAI_API_KEY=${AZURE_OPENAI_API_KEY}


Example 3: Modifying Settings
──────────────────────────────
# In config.py
class ExcelExportConfig:
    HEADER_COLOR = "FF0000"  # Change to red


Configuration Files Overview
═════════════════════════════

┌────────────────────────────────────────────────────────┐
│ File                        │ Purpose                  │
├────────────────────────────────────────────────────────┤
│ config.py                   │ Main configuration       │
│ .env                        │ Secrets & credentials    │
│ .env.example                │ Template for .env        │
│ CONFIGURATION.md            │ Detailed documentation   │
│ CONFIG_QUICK_REFERENCE.md   │ Quick lookup guide       │
│ CONFIGURATION_SUMMARY.md    │ Overview & summary       │
│ CONFIGURATION_DIAGRAM.md    │ This file (visual guide) │
└────────────────────────────────────────────────────────┘


Security Model
══════════════

┌─────────────────────────────────────────────────────────┐
│                    SECURITY LAYERS                       │
└─────────────────────────────────────────────────────────┘

Layer 1: .gitignore
    ↓ (Prevents committing secrets)
Layer 2: .env file
    ↓ (Stores secrets locally)
Layer 3: Environment variables
    ↓ (Runtime injection)
Layer 4: Azure Key Vault (Production)
    ↓ (Centralized secret management)
Application


Best Practices Workflow
════════════════════════

1. Development
   ├── Copy .env.example → .env
   ├── Edit .env with dev credentials
   └── Modify config.py for app settings

2. Testing
   ├── Use separate .env for test environment
   └── Verify configuration with: python config.py

3. Staging
   ├── Use environment-specific .env
   └── Test with production-like settings

4. Production
   ├── Use Azure Key Vault or similar
   ├── Set environment variables in deployment
   └── Never commit .env to repository


Configuration Change Process
═════════════════════════════

┌─────────────────────────────────────────────────────┐
│ Need to change a setting?                           │
└─────────────────────────────────────────────────────┘
                    │
                    ▼
        ┌───────────────────────┐
        │ Is it a secret/key?   │
        └───────────────────────┘
                │           │
            Yes │           │ No
                │           │
                ▼           ▼
        ┌──────────┐   ┌──────────┐
        │ Edit     │   │ Edit     │
        │ .env     │   │ config.py│
        └──────────┘   └──────────┘
                │           │
                └─────┬─────┘
                      │
                      ▼
        ┌───────────────────────┐
        │ Restart application   │
        └───────────────────────┘
                      │
                      ▼
        ┌───────────────────────┐
        │ Test the change       │
        └───────────────────────┘
                      │
                      ▼
        ┌───────────────────────┐
        │ Document if needed    │
        └───────────────────────┘


Quick Reference: Where to Find Settings
════════════════════════════════════════

Azure OpenAI Settings     → .env + config.py (AzureOpenAIConfig)
Server Port              → .env + config.py (AppConfig)
Excel Colors             → config.py (ExcelExportConfig)
Quota Providers          → config.py (AzureQuotaConfig)
Resource Names           → config.py (ResourceTypeFriendlyNames)
Batch Sizes              → config.py (AzureResourceGraphConfig)
AI Analysis Limits       → config.py (AIAnalysisConfig)
Progress Tracking        → config.py (AssessmentProgressConfig)
Network Commands         → config.py (NetworkConfig)
```
