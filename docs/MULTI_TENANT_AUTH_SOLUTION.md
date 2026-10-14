# Multi-Tenant Authentication Solution

## Problem
Azure CLI's `az account list --all` only returns subscriptions from the **home tenant** when using device code authentication. Guest tenant subscriptions are not enumerated reliably.

## Root Cause
When you authenticate with `az login --use-device-code` without specifying a tenant, Azure CLI authenticates to your **home tenant** only. The `--all` flag doesn't actually query all tenants you have access to - it only shows subscriptions within the current authenticated tenant context.

## Solution Implemented
Added support for **tenant-specific authentication** via the `tenant_id` parameter in `/api/auth/device-code-start`.

### Backend Changes
- Modified `start_device_code_login()` to accept optional `tenant_id` parameter
- When `tenant_id` is provided, uses `az login --use-device-code --tenant <tenant_id>`
- This authenticates directly to the specified tenant

### Recommended UX Flow

#### Option 1: Two-Step Authentication (Recommended)
1. **Initial Auth**: User authenticates to home tenant (no tenant_id)
2. **Tenant Selection**: Show list of tenants from `az account tenant list`
3. **Re-Auth for Guest Tenant**: If user selects a guest tenant, trigger new device code flow with that `tenant_id`
4. **Subscription Selection**: Now `az account list` shows subscriptions from the selected tenant

#### Option 2: Pre-Select Tenant
1. **Tenant Input**: Ask user to enter tenant ID/name before authentication
2. **Direct Auth**: Authenticate directly to that tenant
3. **Subscription Selection**: Show subscriptions from that tenant only

## API Usage

### Authenticate to Home Tenant
```javascript
POST /api/auth/device-code-start
// No body needed
```

### Authenticate to Specific Tenant
```javascript
POST /api/auth/device-code-start?tenant_id=6a294346-3fb5-46dd-8278-249322318b53
```

## Benefits
- ✅ **Reliable**: Always shows subscriptions from the authenticated tenant
- ✅ **Complete**: Can access ALL tenants, not just those that happen to work with `--all`
- ✅ **Explicit**: User knows exactly which tenant they're authenticated to
- ✅ **No Azure CLI bugs**: Doesn't rely on unreliable multi-tenant caching

## Trade-offs
- ⚠️ Requires **two device code flows** for guest tenants (extra user interaction)
- ⚠️ User must complete authentication twice if accessing multiple tenants
- ✅ But this is the **only reliable way** to access guest tenant subscriptions via Azure CLI

## Implementation Status
- ✅ Backend API updated to support `tenant_id` parameter
- ⏳ Frontend needs update to implement two-step auth flow
- ⏳ UI needs "Re-authenticate for this tenant" button for guest tenants

## Alternative Considered
Using `az account list --all --refresh` with cache clearing - **Does not work reliably** due to Azure CLI limitations.
