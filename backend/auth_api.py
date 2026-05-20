"""
FastAPI endpoints for Azure Device Code authentication.
Provides web UI integration for authentication flow.
"""

from fastapi import APIRouter, HTTPException, BackgroundTasks
from pydantic import BaseModel
from typing import Optional, Dict, List
from azure.identity import DeviceCodeCredential
from azure.mgmt.subscription import SubscriptionClient
import asyncio
import time
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
import uuid
import subprocess
import threading
import re
import json
import tempfile
import os
import hashlib
import base64
import boto3
import requests

router = APIRouter(prefix="/api/auth", tags=["authentication"])

# In-memory session storage (replace with Redis/database in production)
auth_sessions: Dict[str, dict] = {}

# Device code sessions for az CLI authentication
device_code_sessions: Dict[str, dict] = {}

# GCP OAuth sessions
gcp_oauth_sessions: Dict[str, dict] = {}

# Thread pool for running blocking operations
executor = ThreadPoolExecutor(max_workers=10)


class TenantIdRequest(BaseModel):
    tenant_id: str


class AuthSession(BaseModel):
    session_id: str
    status: str  # "pending", "authenticated", "failed", "expired"
    device_code: Optional[str] = None
    user_code: Optional[str] = None
    verification_url: Optional[str] = None
    message: Optional[str] = None
    expires_at: Optional[str] = None
    error: Optional[str] = None


class SubscriptionInfo(BaseModel):
    subscription_id: str
    display_name: str
    state: str


def cleanup_expired_sessions():
    """Remove expired sessions from memory."""
    now = datetime.now()
    expired = []
    
    for sid, session in auth_sessions.items():
        expires_at_str = session.get("expires_at")
        if expires_at_str:
            try:
                # Handle both string and datetime objects
                if isinstance(expires_at_str, str):
                    expires_at = datetime.fromisoformat(expires_at_str.replace('Z', '+00:00'))
                else:
                    expires_at = expires_at_str
                
                # Make both datetimes timezone-naive for comparison
                if expires_at.tzinfo is not None:
                    expires_at = expires_at.replace(tzinfo=None)
                
                if expires_at < now:
                    expired.append(sid)
            except:
                # If there's any error parsing, skip this session
                pass
    
    for sid in expired:
        del auth_sessions[sid]


def get_device_code_sync(tenant_id: str, device_code_info: dict, session_id: str):
    """
    Synchronous function to get device code.
    This will be run in a thread pool.
    """
    try:
        def device_code_callback(verification_uri, user_code, expires_in):
            device_code_info["verification_url"] = verification_uri
            device_code_info["user_code"] = user_code
            device_code_info["expires_in"] = expires_in
            device_code_info["message"] = f"To sign in, open {verification_uri} and enter the code {user_code}"
            device_code_info["ready"] = True
        
        credential = DeviceCodeCredential(
            tenant_id=tenant_id,
            client_id="04b07795-8ddb-461a-bbee-02f9e1bf7b46",
            prompt_callback=device_code_callback,
            timeout=900
        )
        
        # Store credential for later use
        device_code_info["credential"] = credential
        
        # This will trigger the callback and then wait for authentication
        try:
            credential.get_token("https://management.azure.com/.default")
            # Authentication succeeded
            session = auth_sessions.get(session_id)
            if session:
                session["status"] = "authenticated"
                session["authenticated_at"] = datetime.now().isoformat()
        except Exception as e:
            # Authentication failed or timed out
            session = auth_sessions.get(session_id)
            if session:
                session["status"] = "failed"
                session["error"] = str(e)
                
    except Exception as e:
        device_code_info["error"] = str(e)


@router.post("/start", response_model=AuthSession)
async def start_authentication(request: TenantIdRequest):
    """
    Start Device Code authentication flow.
    Returns device code and verification URL for user to complete authentication.
    """
    cleanup_expired_sessions()
    
    tenant_id = request.tenant_id.strip()
    if not tenant_id:
        raise HTTPException(status_code=400, detail="Tenant ID is required")
    
    session_id = str(uuid.uuid4())
    
    try:
        # Shared dict to capture device code info
        device_code_info = {"ready": False}
        
        # Start the device code flow in a thread
        loop = asyncio.get_event_loop()
        loop.run_in_executor(
            executor,
            get_device_code_sync,
            tenant_id,
            device_code_info,
            session_id
        )
        
        # Wait for device code to be ready (max 10 seconds)
        for _ in range(20):  # 20 * 0.5 = 10 seconds
            await asyncio.sleep(0.5)
            if device_code_info.get("ready"):
                break
        
        if not device_code_info.get("ready"):
            raise HTTPException(status_code=500, detail="Timeout waiting for device code")
        
        if "error" in device_code_info:
            raise HTTPException(status_code=500, detail=device_code_info["error"])
        
        # Store session
        # expires_in from Azure is actually a datetime object, not seconds
        expires_in_value = device_code_info.get("expires_in")
        if isinstance(expires_in_value, datetime):
            expires_at = expires_in_value
        elif isinstance(expires_in_value, (int, float)):
            expires_at = datetime.now() + timedelta(seconds=expires_in_value)
        else:
            # Default to 15 minutes
            expires_at = datetime.now() + timedelta(seconds=900)
            
        auth_sessions[session_id] = {
            "tenant_id": tenant_id,
            "credential": device_code_info.get("credential"),
            "status": "pending",
            "device_code_info": device_code_info,
            "expires_at": expires_at.isoformat(),
            "created_at": datetime.now().isoformat()
        }
        
        return AuthSession(
            session_id=session_id,
            status="pending",
            user_code=device_code_info.get("user_code"),
            verification_url=device_code_info.get("verification_url"),
            message=device_code_info.get("message"),
            expires_at=expires_at.isoformat()
        )
        
    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Failed to start authentication: {str(e)}")


@router.get("/status/{session_id}", response_model=AuthSession)
async def get_auth_status(session_id: str):
    """
    Check the status of an authentication session.
    """
    cleanup_expired_sessions()
    
    session = auth_sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found or expired")
    
    device_code_info = session.get("device_code_info", {})
    
    return AuthSession(
        session_id=session_id,
        status=session["status"],
        user_code=device_code_info.get("user_code"),
        verification_url=device_code_info.get("verification_url"),
        message=device_code_info.get("message"),
        expires_at=session.get("expires_at"),
        error=session.get("error")
    )


@router.get("/subscriptions/{session_id}", response_model=List[SubscriptionInfo])
async def list_subscriptions(session_id: str):
    """
    List accessible Azure subscriptions after successful authentication.
    """
    session = auth_sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    
    if session["status"] != "authenticated":
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    try:
        credential = session["credential"]
        sub_client = SubscriptionClient(credential)
        
        subscriptions = []
        for sub in sub_client.subscriptions.list():
            subscriptions.append(SubscriptionInfo(
                subscription_id=sub.subscription_id,
                display_name=sub.display_name,
                state=sub.state
            ))
        
        return subscriptions
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to list subscriptions: {str(e)}")


@router.delete("/logout/{session_id}")
async def logout(session_id: str):
    """
    End authentication session and clear credentials.
    """
    if session_id in auth_sessions:
        del auth_sessions[session_id]
        return {"message": "Logged out successfully"}
    
    raise HTTPException(status_code=404, detail="Session not found")


# ============================================================================
# Alternative: az CLI-based Device Code Authentication
# ============================================================================

@router.post("/device-code-start")
async def start_device_code_login(tenant_id: Optional[str] = None):
    """
    Start az login --use-device-code and return the device code and URL.
    
    Args:
        tenant_id: Optional tenant ID to authenticate to a specific tenant.
                   If not provided, authenticates to the user's home tenant.
    
    This is an alternative to the SDK-based /start endpoint.
    Uses az CLI commands instead of Azure SDK.
    """
    session_id = str(uuid.uuid4())
    
    try:
        # Create isolated environment for this session
        session_env_dir = tempfile.mkdtemp(prefix=f"az_session_{session_id}_")
        custom_env = os.environ.copy()
        custom_env["AZURE_CONFIG_DIR"] = session_env_dir
        print(f"🔒 Created isolated session env: {session_env_dir}")

        # Build az login command
        login_cmd = ["az", "login", "--use-device-code"]
        if tenant_id:
            login_cmd.extend(["--tenant", tenant_id])
            print(f"🎯 Authenticating to specific tenant: {tenant_id}")
        else:
            print("🚀 Authenticating to home tenant")

        # Start az login --use-device-code in a subprocess with isolated env
        print(f"🚀 Starting: {' '.join(login_cmd)}...")
        process = subprocess.Popen(
            login_cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,  # Redirect stderr to stdout
            text=True,
            bufsize=1,  # Line buffered
            env=custom_env
        )
        
        # Read output to get the device code and URL
        output_lines = []
        user_code = None
        url = None
        
        print("📖 Reading az CLI output...")
        # Read lines until we have both code and URL or timeout
        start_time = time.time()
        while time.time() - start_time < 10:  # 10 second timeout
            line = process.stdout.readline()
            if not line:
                break
                
            line = line.strip()
            if line:
                print(f"  📄 {line}")
                output_lines.append(line)
                
                # Extract URL
                if "https://" in line and not url:
                    urls = re.findall(r'https://[^\s]+', line)
                    if urls:
                        url = urls[0]
                        print(f"  ✅ Found URL: {url}")
                
                # Extract code - look for patterns like "code ABCD1234" or just "ABCD1234"
                if not user_code:
                    # Pattern 1: "code XXXXXXXX"
                    code_match = re.search(r'\bcode\s+([A-Z0-9]{8,})\b', line, re.IGNORECASE)
                    if code_match:
                        user_code = code_match.group(1)
                        print(f"  ✅ Found code (pattern 1): {user_code}")
                    else:
                        # Pattern 2: standalone uppercase alphanumeric code
                        words = line.split()
                        for word in words:
                            # Check if it's an uppercase alphanumeric string of 8+ chars
                            if len(word) >= 8 and word.isupper() and word.isalnum():
                                user_code = word
                                print(f"  ✅ Found code (pattern 2): {user_code}")
                                break
                
                # If we have both, we can stop reading
                if user_code and url:
                    print("  ✅ Got both code and URL!")
                    break
        
        print(f"📋 Final - Code: {user_code}, URL: {url}")
        
        # If we didn't get the code or URL, log the full output for debugging
        if not user_code or not url:
            print("⚠️ Failed to extract code or URL. Full output:")
            for line in output_lines:
                print(f"    {line}")
        
        # Store the process in the session
        device_code_sessions[session_id] = {
            "process": process,
            "status": "pending",
            "user_code": user_code,
            "url": url,
            "created_at": datetime.now().isoformat(),
            "output_lines": output_lines,  # Store for debugging
            "env_dir": session_env_dir
        }
        # Start a thread to wait for completion
        def wait_for_completion():
            process.wait()
            if process.returncode == 0:
                device_code_sessions[session_id]["status"] = "completed"
                device_code_sessions[session_id]["completed_at"] = datetime.now().isoformat()

                try:
                    custom_env_for_tenant = os.environ.copy()
                    custom_env_for_tenant["AZURE_CONFIG_DIR"] = session_env_dir

                    account_result = subprocess.run(
                        ["az", "account", "show", "--output", "json"],
                        capture_output=True,
                        text=True,
                        check=True,
                        env=custom_env_for_tenant
                    )
                    account_info = json.loads(account_result.stdout)
                    device_code_sessions[session_id]["tenant_id"] = account_info.get("tenantId")
                    device_code_sessions[session_id]["tenant_name"] = account_info.get("name", "Unknown")
                    print(f"Authentication completed for session {session_id} - Tenant: {account_info.get('name')} ({account_info.get('tenantId')})")
                except Exception as e:
                    print(f"Could not fetch tenant info: {e}")
                    print(f"Authentication completed for session {session_id}")

            else:
                device_code_sessions[session_id]["status"] = "failed"
                remaining = process.stdout.read()
                device_code_sessions[session_id]["error"] = remaining
                print(f"Authentication failed for session {session_id}: {remaining}")

        thread = threading.Thread(target=wait_for_completion)
        thread.daemon = True
        thread.start()
        
        return {
            "session_id": session_id,
            "user_code": user_code,
            "verification_url": url or "https://microsoft.com/devicelogin",
            "message": "Please visit the URL and enter the code to authenticate"
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to start device code login: {str(e)}")


@router.get("/device-code-status/{session_id}")
async def check_device_code_status(session_id: str):
    """
    Check the status of a device code authentication session.
    Works with sessions created by /device-code-start endpoint.
    """
    if session_id not in device_code_sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    session = device_code_sessions[session_id]

    # Fastest completion signal: if az login process already exited with success,
    # mark session completed immediately so UI can move on.
    if session.get("status") == "pending" and session.get("process") is not None:
        try:
            process = session["process"]
            return_code = process.poll()
            if return_code == 0:
                session["status"] = "completed"
                session["completed_at"] = datetime.now().isoformat()
            elif return_code is not None and return_code != 0:
                session["status"] = "failed"
                if not session.get("error"):
                    session["error"] = "Authentication process exited with non-zero status."
        except Exception:
            pass

    # Fast-path completion detection:
    # Do not wait only for the background az login process watcher.
    # If current session can run 'az account show', treat auth as completed immediately.
    if session.get("status") == "pending":
        last_probe_at = session.get("last_probe_at")
        should_probe = True

        if last_probe_at:
            try:
                last_probe_dt = datetime.fromisoformat(last_probe_at)
                should_probe = (datetime.now() - last_probe_dt).total_seconds() >= 1
            except Exception:
                should_probe = True

        if should_probe:
            session["last_probe_at"] = datetime.now().isoformat()
            try:
                custom_env = os.environ.copy()
                if "env_dir" in session:
                    custom_env["AZURE_CONFIG_DIR"] = session["env_dir"]

                account_result = subprocess.run(
                    ["az", "account", "show", "--output", "json"],
                    capture_output=True,
                    text=True,
                    env=custom_env,
                    timeout=2
                )

                if account_result.returncode == 0 and account_result.stdout:
                    account_info = json.loads(account_result.stdout)
                    session["status"] = "completed"
                    session["completed_at"] = datetime.now().isoformat()
                    session["tenant_id"] = account_info.get("tenantId")
                    session["tenant_name"] = account_info.get("name", "Unknown")
            except Exception:
                # Keep pending if probe fails; background watcher still updates status.
                pass
    
    return {
        "status": session["status"],
        "user_code": session.get("user_code"),
        "verification_url": session.get("url"),
        "error": session.get("error"),
        "tenant_id": session.get("tenant_id"),
        "tenant_name": session.get("tenant_name")
    }


@router.get("/device-code-tenants/{session_id}")
async def list_cli_tenants(session_id: str, probe_missing_tenants: bool = False):
    """
    List distinct tenants available in the authenticated session.

    This uses a multi-step approach:
    1. Get all tenants from 'az account tenant list'
    2. For each tenant, try to list subscriptions
    3. Only return tenants that have accessible subscriptions
    """
    if session_id not in device_code_sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = device_code_sessions[session_id]

    if session["status"] != "completed":
        raise HTTPException(status_code=401, detail="Session not authenticated")

    try:
        custom_env = os.environ.copy()
        if "env_dir" in session:
            custom_env["AZURE_CONFIG_DIR"] = session["env_dir"]
            
        # Step 1: Get all tenants the user has access to
        print("Fetching tenant list...")
        tenant_result = subprocess.run(
            ["az", "account", "tenant", "list", "--output", "json"],
            capture_output=True, text=True, env=custom_env
        )
        
        all_tenants = []
        if tenant_result.returncode == 0:
            try:
                all_tenants = json.loads(tenant_result.stdout)
            except:
                pass
        
        # Step 2: Get all subscriptions from 'az account list --all'
        result = subprocess.run(
            ["az", "account", "list", "--all", "--output", "json"],
            capture_output=True, text=True, env=custom_env
        )
        
        subscriptions_data = []
        if result.returncode == 0 and result.stdout:
            try:
                subscriptions_data = json.loads(result.stdout)
            except Exception:
                pass
        
        # Step 3: Build tenant map with subscription counts
        tenants = {}
        for sub in subscriptions_data:
            t_id = sub.get("tenantId")
            if t_id:
                if t_id not in tenants:
                    tenants[t_id] = {
                        "tenant_id": t_id,
                        "display_name": "Unknown Tenant",
                        "subscription_count": 0
                    }
                tenants[t_id]["subscription_count"] += 1
        
        # Step 4: Enrich with display names from tenant list
        for t in all_tenants:
            t_id = t.get("tenantId")
            if t_id in tenants:
                tenants[t_id]["display_name"] = t.get("displayName", t_id)
        
        # Step 5: Include tenants with zero subscriptions (and optional explicit probe)
        for t in all_tenants:
            t_id = t.get("tenantId")
            if t_id and t_id not in tenants:
                name = t.get("displayName", t_id)
                tenants[t_id] = {
                    "tenant_id": t_id,
                    "display_name": name if name else f"Tenant {t_id[:8]}...",
                    "subscription_count": 0
                }
                
                if probe_missing_tenants:
                    try:
                        tenant_sub_result = subprocess.run(
                            ["az", "account", "list", "--query", f"[?tenantId=='{t_id}']", "--output", "json"],
                            capture_output=True, text=True, env=custom_env, timeout=10
                        )
                        
                        if tenant_sub_result.returncode == 0 and tenant_sub_result.stdout:
                            try:
                                tenant_subs = json.loads(tenant_sub_result.stdout)
                                if len(tenant_subs) > 0:
                                    tenants[t_id]["subscription_count"] = len(tenant_subs)
                            except:
                                pass
                    except Exception:
                        pass
        
        for t_id, t_data in tenants.items():
            if t_data["display_name"] == "Unknown Tenant":
                t_data["display_name"] = f"Tenant {t_id[:8]}..."

        tenant_list = list(tenants.values())
        return tenant_list

    except Exception as e:
        print(f"Error getting tenants: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Failed to list tenants: {str(e)}")


@router.get("/device-code-subscriptions/{session_id}")
async def list_cli_subscriptions(session_id: str, tenant_id: Optional[str] = None):
    """
    List Azure subscriptions for a CLI-authenticated session.
    Uses az account list command.
    Optional: tenant_id to filter subscriptions by tenant.
    """
    print(f"📥 list_cli_subscriptions called. Session: {session_id}, Tenant Filter: {tenant_id}")
    
    if session_id not in device_code_sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    session = device_code_sessions[session_id]
    
    if session["status"] != "completed":
        raise HTTPException(status_code=401, detail="Session not authenticated")
    
    try:
        # Use isolated environment if available
        custom_env = os.environ.copy()
        if "env_dir" in session:
            custom_env["AZURE_CONFIG_DIR"] = session["env_dir"]

        # Get current account info (best effort)
        current_tenant_id = None
        try:
            account_result = subprocess.run(
                ["az", "account", "show", "--output", "json"],
                capture_output=True, text=True, env=custom_env
            )
            if account_result.returncode == 0:
                account_info = json.loads(account_result.stdout)
                current_tenant_id = account_info.get("tenantId")
        except Exception as e:
            print(f"⚠️ Failed to get current tenant (ignored): {e}")

        # Use az account list --all to get all subscriptions
        print(f"📋 Running az account list --all --refresh...")
        result = subprocess.run(
            ["az", "account", "list", "--all", "--refresh", "--output", "json"],
            capture_output=True,
            text=True,
            env=custom_env
        )
        
        if result.returncode != 0:
            print(f"⚠️ az account list returned non-zero exit code: {result.returncode}")
            print(f"Stderr: {result.stderr}")
        
        if not result.stdout or result.stdout.strip() == "":
            print(f"⚠️ az account list returned empty output")
            subscriptions_data = []
        else:
            try:
                subscriptions_data = json.loads(result.stdout)
            except json.JSONDecodeError as e:
                print(f"❌ Failed to parse az account list output: {e}")
                print(f"Output was: {result.stdout[:200]}...")
                raise HTTPException(status_code=500, detail=f"Invalid JSON from Azure CLI: {str(e)}")
        
        # Log all subscriptions with their tenant IDs for debugging
        print(f"🔍 Raw subscriptions from CLI ({len(subscriptions_data)} total):")
        tenant_summary = {}
        for sub in subscriptions_data:
            t_id = sub.get("tenantId", "UNKNOWN")
            tenant_summary[t_id] = tenant_summary.get(t_id, 0) + 1
        for t_id, count in tenant_summary.items():
            print(f"   Tenant {t_id}: {count} subscription(s)")
        
        # Determine target tenant
        target_tenant = tenant_id if tenant_id else current_tenant_id
        
        print(f"🎯 Filtering for Tenant: {target_tenant} (Requested: {tenant_id}, Current: {current_tenant_id})")

        # Filter subscriptions
        subscriptions = []
        filtered_count = 0
        
        for sub in subscriptions_data:
            sub_tenant_id = sub.get("tenantId")
            
            # If target_tenant is specified, filter by it.
            if target_tenant and sub_tenant_id != target_tenant:
                filtered_count += 1
                continue
            
            subscriptions.append({
                "subscription_id": sub.get("id"),
                "display_name": sub.get("name"),
                "state": sub.get("state"),
                "is_default": sub.get("isDefault", False),
                "tenant_id": sub_tenant_id
            })
        
        if filtered_count > 0:
            print(f"ℹ️  Filtered out {filtered_count} subscription(s) (Target: {target_tenant})")
        
        print(f"✅ Returning {len(subscriptions)} subscription(s)")
        return subscriptions
        
    except Exception as e:
        print(f"❌ Unexpected Error in list_cli_subscriptions: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Internal Server Error: {str(e)}")


# ============================================================================
# AWS Authentication
# ============================================================================

class AwsCredentials(BaseModel):
    aws_access_key_id: str
    aws_secret_access_key: str
    aws_session_token: Optional[str] = None
    region_name: str

@router.post("/aws/validate")
async def validate_aws_credentials(creds: AwsCredentials):
    """
    Validate AWS credentials by calling STS GetCallerIdentity.
    """
    try:
        session = boto3.Session(
            aws_access_key_id=creds.aws_access_key_id,
            aws_secret_access_key=creds.aws_secret_access_key,
            aws_session_token=creds.aws_session_token,
            region_name=creds.region_name
        )
        sts = session.client('sts')
        identity = sts.get_caller_identity()
        
        return {
            "status": "valid",
            "account_id": identity['Account'],
            "arn": identity['Arn'],
            "user_id": identity['UserId']
        }
    except Exception as e:
        print(f"❌ AWS Auth Failed: {str(e)}")
        raise HTTPException(status_code=401, detail=f"AWS Authentication failed: {str(e)}")


@router.post("/aws/discover-regions")
async def discover_aws_regions(creds: AwsCredentials):
    """
    Discover which AWS regions have active resources.
    Parallelizes checks across all enabled regions.
    """
    try:
        # 1. Use STS to get account access and list enabled regions
        # We try the user-provided region first, then fallback to us-east-1
        base_region = creds.region_name or 'us-east-1'
        
        def get_all_regions(region):
            session = boto3.Session(
                aws_access_key_id=creds.aws_access_key_id,
                aws_secret_access_key=creds.aws_secret_access_key,
                aws_session_token=creds.aws_session_token,
                region_name=region
            )
            ec2 = session.client('ec2')
            resp = ec2.describe_regions()
            return [r['RegionName'] for r in resp['Regions']]

        try:
            all_regions = get_all_regions(base_region)
        except Exception as e:
            if base_region != 'us-east-1':
                print(f"⚠️ Initial discovery failed in {base_region}, retrying in us-east-1...")
                all_regions = get_all_regions('us-east-1')
            else:
                raise e
        
        print(f"🔍 Discovering resources across {len(all_regions)} regions...")
        active_regions = []
        
        def check_region(region_name):
            try:
                # Use a session per thread for safety
                session = boto3.Session(
                    aws_access_key_id=creds.aws_access_key_id,
                    aws_secret_access_key=creds.aws_secret_access_key,
                    aws_session_token=creds.aws_session_token,
                    region_name=region_name
                )
                
                has_res = False
                
                # Check 1: Tagging API (Fastest & Broadest)
                try:
                    tagging = session.client('resourcegroupstaggingapi')
                    resp = tagging.get_resources(ResourcesPerPage=1)
                    if resp.get('ResourceTagMappingList'):
                        has_res = True
                except Exception:
                    pass # Continue to direct services if Denied
                
                # Check 2: EC2 Instances (Running/Stopped)
                if not has_res:
                    try:
                        ec2 = session.client('ec2')
                        resp = ec2.describe_instances(MaxResults=5)
                        if resp.get('Reservations'):
                            has_res = True
                    except Exception:
                        pass

                # Check 3: RDS Instances
                if not has_res:
                    try:
                        rds = session.client('rds')
                        resp = rds.describe_db_instances(MaxRecords=1, # RDS requires 1-100
                                                       # If it fails with "must be 20", it will be caught
                        )
                        if resp.get('DBInstances'):
                            has_res = True
                    except Exception as rds_e:
                        # Fallback for some accounts that enforce minimum count
                        if "must be between" in str(rds_e):
                            try:
                                resp = rds.describe_db_instances(MaxRecords=20)
                                if resp.get('DBInstances'):
                                    has_res = True
                            except: pass

                # Check 4: Network Infrastructure (Proxy for any VPC activity)
                # We check for non-default Security Groups or non-default VPCs
                if not has_res:
                    try:
                        ec2 = session.client('ec2')
                        # Security groups are always there, but if there's more than 1, it's usually user-created
                        sgs = ec2.describe_security_groups(MaxResults=5)
                        if len(sgs.get('SecurityGroups', [])) > 1:
                            has_res = True
                        
                        if not has_res:
                            vpcs = ec2.describe_vpcs(MaxResults=5)
                            # If there's more than 1 VPC, it's definitely an active region
                            if len(vpcs.get('Vpcs', [])) > 1:
                                has_res = True
                            # Even with 1 VPC, if it's not the default one, it's active
                            elif vpcs.get('Vpcs') and not vpcs['Vpcs'][0].get('IsDefault', True):
                                has_res = True
                    except Exception:
                        pass
                
                # Check 5: Lambda Functions
                if not has_res:
                    try:
                        lmb = session.client('lambda')
                        resp = lmb.list_functions(MaxItems=1)
                        if resp.get('Functions'):
                            has_res = True
                    except Exception:
                        pass

                if has_res:
                    return {
                        "id": region_name,
                        "has_resources": True
                    }
                return None
            except Exception as e:
                # Log only unexpected server/network errors
                low_msg = str(e).lower()
                if "unavailable" in low_msg or "timeout" in low_msg or "connection" in low_msg:
                     print(f"⚠️  Region {region_name} unreachable: {str(e)}")
                return None

        # Execute discovery in parallel
        with ThreadPoolExecutor(max_workers=25) as r_executor:
            future_to_region = {r_executor.submit(check_region, r): r for r in all_regions}
            for future in future_to_region:
                result = future.result()
                if result:
                    active_regions.append(result)
        
        active_regions.sort(key=lambda x: x['id'])
        found_names = [r['id'] for r in active_regions]
        print(f"✅ Discovery complete. Found {len(active_regions)} active regions: {', '.join(found_names)}")
        return active_regions

    except Exception as e:
        err_msg = str(e)
        if "AuthFailure" in err_msg or "InvalidClientTokenId" in err_msg or "SignatureDoesNotMatch" in err_msg:
             print(f"❌ AWS Credential Failure: {err_msg}")
             raise HTTPException(status_code=401, detail="AWS was not able to validate the provided access credentials. Please check your Access Key and Secret Key.")
        
        print(f"❌ AWS Region Discovery Failed: {err_msg}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Failed to discover regions: {err_msg}")


# ============================================================================
# GCP Authentication
# ============================================================================

class GcpOauthStartRequest(BaseModel):
    project_id: Optional[str] = None


class GcpOauthExchangeRequest(BaseModel):
    session_id: str
    state: str
    code: str


def _urlsafe_b64encode(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("utf-8").rstrip("=")


def _cleanup_expired_gcp_oauth_sessions():
    now = datetime.utcnow()
    expired_session_ids = []
    for sid, session in gcp_oauth_sessions.items():
        expires_at = session.get("expires_at")
        if isinstance(expires_at, datetime) and expires_at < now:
            expired_session_ids.append(sid)
    for sid in expired_session_ids:
        del gcp_oauth_sessions[sid]


@router.post("/gcp/oauth-start")
async def start_gcp_oauth(request: GcpOauthStartRequest):
    """
    Start GCP OAuth authorization code flow with PKCE.
    Frontend opens returned auth_url in a popup, then calls /gcp/oauth-exchange.
    """
    _cleanup_expired_gcp_oauth_sessions()

    project_id = (request.project_id or "").strip()
    # project_id is now optional for start

    client_id = os.getenv("GCP_OAUTH_CLIENT_ID", "").strip()
    client_secret = os.getenv("GCP_OAUTH_CLIENT_SECRET", "").strip()
    redirect_uri = os.getenv("GCP_OAUTH_REDIRECT_URI", "").strip()

    if not client_id:
        raise HTTPException(
            status_code=500,
            detail="GCP OAuth is not configured. Set GCP_OAUTH_CLIENT_ID and GCP_OAUTH_REDIRECT_URI."
        )
    if not redirect_uri:
        raise HTTPException(
            status_code=500,
            detail="GCP OAuth is not configured. Set GCP_OAUTH_REDIRECT_URI."
        )

    session_id = str(uuid.uuid4())
    state = session_id
    code_verifier = _urlsafe_b64encode(os.urandom(48))
    code_challenge = _urlsafe_b64encode(hashlib.sha256(code_verifier.encode("utf-8")).digest())

    scope = " ".join([
        "openid",
        "email",
        "profile",
        "https://www.googleapis.com/auth/cloud-platform.read-only",
    ])

    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": scope,
        "state": state,
        "access_type": "offline",
        "prompt": "consent",
        "include_granted_scopes": "true",
        "code_challenge": code_challenge,
        "code_challenge_method": "S256",
    }
    auth_url = "https://accounts.google.com/o/oauth2/v2/auth?" + "&".join(
        [f"{k}={requests.utils.quote(str(v), safe='')}" for k, v in params.items()]
    )

    gcp_oauth_sessions[session_id] = {
        "session_id": session_id,
        "status": "pending",
        "project_id": project_id,
        "state": state,
        "code_verifier": code_verifier,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "created_at": datetime.utcnow(),
        "expires_at": datetime.utcnow() + timedelta(minutes=15),
    }

    return {
        "session_id": session_id,
        "auth_url": auth_url,
        "project_id": project_id,
        "expires_in_seconds": 900,
    }


@router.post("/gcp/oauth-exchange")
async def exchange_gcp_oauth_code(request: GcpOauthExchangeRequest):
    """
    Exchange authorization code for access token and validate project access.
    """
    session = gcp_oauth_sessions.get(request.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="OAuth session not found or expired")

    if session.get("status") == "completed":
        return {
            "status": "valid",
            "project_id": session.get("project_id"),
            "project_name": session.get("project_name", session.get("project_id")),
            "session_id": request.session_id,
        }

    if request.state != session.get("state"):
        raise HTTPException(status_code=400, detail="Invalid OAuth state")

    try:
        token_resp = requests.post(
            "https://oauth2.googleapis.com/token",
            data={
                "code": request.code,
                "client_id": session["client_id"],
                "client_secret": session["client_secret"],
                "redirect_uri": session["redirect_uri"],
                "grant_type": "authorization_code",
                "code_verifier": session["code_verifier"],
            },
            timeout=30,
        )
        if token_resp.status_code != 200:
            raise HTTPException(
                status_code=401,
                detail=f"Failed to exchange OAuth code: {token_resp.text}"
            )

        token_data = token_resp.json()
        access_token = token_data.get("access_token")
        if not access_token:
            raise HTTPException(status_code=401, detail="No access token returned by Google")

        session["status"] = "completed"
        session["access_token"] = access_token
        session["refresh_token"] = token_data.get("refresh_token")
        session["scope"] = token_data.get("scope")
        session["token_type"] = token_data.get("token_type")
        session["verified_at"] = datetime.utcnow()
        session["expires_at"] = datetime.utcnow() + timedelta(seconds=int(token_data.get("expires_in", 3600)))

        project_id = session.get("project_id")
        if project_id:
            project_resp = requests.get(
                f"https://cloudresourcemanager.googleapis.com/v1/projects/{project_id}",
                headers={"Authorization": f"Bearer {access_token}"},
                timeout=30,
            )
            if project_resp.status_code == 200:
                project_info = project_resp.json()
                session["project_name"] = project_info.get("name", project_id)
            else:
                # If we have a project_id but can't access it, we might want to warn or let it fail later
                print(f"⚠️ Warning: Could not verify project {project_id} during exchange")

        return {
            "status": "valid",
            "project_id": project_id,
            "project_name": session["project_name"],
            "session_id": request.session_id,
        }
    except HTTPException:
        raise
    except Exception as e:
        print(f"❌ GCP OAuth exchange failed: {str(e)}")
        raise HTTPException(status_code=500, detail=f"GCP OAuth exchange failed: {str(e)}")


@router.get("/gcp/oauth-status/{session_id}")
async def get_gcp_oauth_status(session_id: str):
    session = gcp_oauth_sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="OAuth session not found")

    response = {
        "session_id": session_id,
        "status": session.get("status", "pending"),
        "project_id": session.get("project_id"),
    }
    if session.get("status") == "completed":
        response["project_name"] = session.get("project_name")
    return response


@router.get("/gcp/projects/{session_id}")
async def list_gcp_projects(session_id: str):
    """
    List accessible GCP projects for an authenticated OAuth session.
    """
    session = gcp_oauth_sessions.get(session_id)
    if not session or session.get("status") != "completed":
        raise HTTPException(status_code=401, detail="Session not authenticated")

    access_token = session.get("access_token")
    try:
        projects = []
        page_token = None
        
        while True:
            params = {}
            if page_token:
                params["pageToken"] = page_token
                
            resp = requests.get(
                "https://cloudresourcemanager.googleapis.com/v1/projects",
                headers={"Authorization": f"Bearer {access_token}"},
                params=params,
                timeout=30
            )
            
            if resp.status_code != 200:
                print(f"⚠️ GCP Project list error ({resp.status_code}): {resp.text}")
                # If we have some projects already, return them; otherwise, fail.
                if projects:
                    break
                raise HTTPException(status_code=resp.status_code, detail=f"Failed to list GCP projects: {resp.text}")
            
            data = resp.json()
            batch = data.get("projects", [])
            projects.extend(batch)
            
            page_token = data.get("nextPageToken")
            if not page_token:
                break
        
        # Filter and normalize
        result = [
            {
                "project_id": p.get("projectId"),
                "display_name": p.get("name", p.get("projectId")),
                "project_number": p.get("projectNumber")
            }
            for p in projects
            if p.get("lifecycleState") == "ACTIVE"
        ]
        
        # Sort by display name
        result.sort(key=lambda x: x["display_name"].lower())
        
        return result
    except Exception as e:
        print(f"❌ Failed to list GCP projects: {str(e)}")
        if isinstance(e, HTTPException):
            raise
        raise HTTPException(status_code=500, detail=f"Failed to list GCP projects: {str(e)}")


class GcpCredentials(BaseModel):
    project_id: Optional[str] = None
    project_ids: Optional[List[str]] = None
    credentials_json: str

@router.post("/gcp/validate")
async def validate_gcp_credentials(creds: GcpCredentials):
    """
    Validate GCP credentials by attempting to access the project(s).
    """
    try:
        from google.oauth2 import service_account
        from google.cloud import resourcemanager_v3
        import json as json_lib
        
        # Parse credentials JSON
        try:
            credentials_dict = json_lib.loads(creds.credentials_json)
        except json_lib.JSONDecodeError as e:
            raise HTTPException(status_code=400, detail=f"Invalid JSON format: {str(e)}")
        
        # Create credentials object
        credentials = service_account.Credentials.from_service_account_info(credentials_dict)
        
        # Validate by trying to access the project(s)
        client = resourcemanager_v3.ProjectsClient(credentials=credentials)
        
        # Determine project list
        project_ids = []
        if creds.project_ids:
            project_ids = creds.project_ids
        elif creds.project_id:
            project_ids = [creds.project_id]
        
        if not project_ids:
            raise HTTPException(status_code=400, detail="No project IDs provided.")

        validated_projects = []
        for pid in project_ids:
            # Heuristic: If user accidentally provided Service Account email instead of Project ID
            if "@" in pid and ".iam.gserviceaccount.com" in pid:
                parts = pid.split("@")
                if len(parts) > 1:
                    domain_parts = parts[1].split(".")
                    if domain_parts:
                        extracted_pid = domain_parts[0]
                        print(f"🔄 Normalizing Project ID from service account email: {pid} -> {extracted_pid}")
                        pid = extracted_pid

            try:
                project_name = f"projects/{pid}"
                project = client.get_project(name=project_name)
                validated_projects.append({
                    "id": pid,
                    "name": project.display_name,
                    "valid": True
                })
            except Exception as e:
                # Fallback: try search
                try:
                    projects = list(client.search_projects(query=f"id:{pid}"))
                    if projects:
                        validated_projects.append({
                            "id": pid,
                            "name": projects[0].display_name,
                            "valid": True
                        })
                    else:
                        raise Exception(f"Project '{pid}' not found.")
                except Exception as inner:
                    # Fallback 2: Try Cloud Asset API (common for Cloud Asset Viewer role)
                    try:
                        from google.cloud import asset_v1
                        asset_client = asset_v1.AssetServiceClient(credentials=credentials)
                        parent = f"projects/{pid}"
                        # Try to list just one asset to verify access
                        response = asset_client.list_assets(request={"parent": parent, "page_size": 1})
                        # Trigger the generator to see if it works
                        next(iter(response), None)
                        
                        validated_projects.append({
                            "id": pid,
                            "name": pid, # Fallback to ID if display name is inaccessible
                            "valid": True,
                            "note": "Validated via Cloud Asset API"
                        })
                    except Exception as asset_err:
                        print(f"⚠️ Validation failed for {pid}: RM Error: {inner}, Asset Error: {asset_err}")
                        raise HTTPException(
                            status_code=403,
                            detail=f"Cannot access project '{pid}'. Please ensure the Service Account has 'Cloud Asset Viewer' or 'Browser/Viewer' roles. Technical error: {str(asset_err)}"
                        )

                    except Exception as asset_err:
                        err_msg = str(asset_err)
                        print(f"⚠️ Validation failed for {pid}: RM Error: {inner}, Asset Error: {asset_err}")
                        
                        # Check for common "API not enabled" error
                        if "is not enabled" in err_msg or "403" in err_msg:
                            raise HTTPException(
                                status_code=403,
                                detail=f"Cannot access project '{pid}'. Please ensure the 'Cloud Asset API' is enabled in the Google Cloud Console for this project and the Service Account has 'Cloud Asset Viewer' role."
                            )
                            
                        raise HTTPException(
                            status_code=403,
                            detail=f"Cannot access project '{pid}': {err_msg}"
                        )

        # Re-build project_ids from validated projects to use normalized IDs
        final_project_ids = [p["id"] for p in validated_projects]

        return {
            "status": "valid",
            "project_ids": final_project_ids,
            "validated_projects": validated_projects,
            "service_account": credentials_dict.get("client_email", "Unknown")
        }
                
    except HTTPException:
        raise
    except Exception as e:
        print(f"❌ GCP Auth Failed: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=401, detail=f"GCP Authentication failed: {str(e)}")



