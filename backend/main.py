from fastapi import FastAPI, HTTPException, BackgroundTasks, Body, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import subprocess
import sys
import os
import shutil
import uuid
import tempfile
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import openai
import pandas as pd
from typing import Optional
from datetime import datetime
import logging

# Suppress frequent polling logs for assessment status
class EndpointFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        return record.getMessage().find("/api/assessment-status/") == -1

# Filter out /api/assessment-status from logging
logging.getLogger("uvicorn.access").addFilter(EndpointFilter())

# Add parent directory to path to import script.py and config
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import script
import aws_assessment
import architecture_generator

# Import centralized configuration
from config import (
    AppConfig,
    AzureOpenAIConfig,
    AIAnalysisConfig,
    AssessmentProgressConfig,
    ExcelExportConfig
)

# Import authentication router and sessions
from auth_api import router as auth_router, device_code_sessions, gcp_oauth_sessions

app = FastAPI()

# Include authentication router
app.include_router(auth_router)

# In-memory job store for async assessments
assessment_jobs = {}
assessment_jobs_lock = threading.Lock()

class LoginRequest(BaseModel):
    auth_type: str = "service_principal"  # service_principal, access_token
    tenant_id: Optional[str] = None
    client_id: Optional[str] = None
    client_secret: Optional[str] = None
    access_token: Optional[str] = None

class ArchitectureRequest(BaseModel):
    file_path: str
    diagram_type: str = "network"
    cloud_type: Optional[str] = None

class AIAssessmentRequest(BaseModel):
    file_path: str
    subscription_id: Optional[str] = None
    migration_type: str = "azure_to_azure"
    access_token: Optional[str] = None  # Required for isolated sessions
    session_id: Optional[str] = None  # For device code sessions

class AssessmentRequest(BaseModel):
    subscription_id: Optional[str] = None
    migration_type: str = "azure_to_azure"
    access_token: Optional[str] = None  # Required for isolated sessions
    session_id: Optional[str] = None  # For device code sessions

@app.get("/health")
async def health_check():
    return {"status": "ok", "message": "Service is healthy"}

@app.post("/api/login")
def login(request: LoginRequest):
    # For Service Principal, we still do a global login (backward compatibility/admin use)
    if request.auth_type == "service_principal":
        try:
            cmd = [
                "az", "login", "--service-principal",
                "-u", request.client_id,
                "-p", request.client_secret,
                "--tenant", request.tenant_id
            ]
            subprocess.run(cmd, capture_output=True, text=True, check=True)
            return {"status": "success", "message": "Logged in successfully (Global Session)"}
        except subprocess.CalledProcessError as e:
            raise HTTPException(status_code=401, detail=f"Login failed: {e.stderr}")
    
    elif request.auth_type == "access_token":
        # For access token, we don't do anything here. 
        # The token will be passed to /api/assess for per-request login.
        return {"status": "success", "message": "Token received. Ready to assess."}
    
    else:
        raise HTTPException(status_code=400, detail="Invalid auth_type")

@app.post("/api/assess")
def run_assessment(request: AssessmentRequest, background_tasks: BackgroundTasks):
    # Determine environment to use
    temp_dir = None
    custom_env = os.environ.copy()
    
    try:
        if request.session_id and request.session_id in device_code_sessions:
            session = device_code_sessions[request.session_id]
            if "env_dir" in session:
                print(f"🔒 Using existing isolated session env: {session['env_dir']}")
                custom_env["AZURE_CONFIG_DIR"] = session["env_dir"]
            else:
                print("⚠️ Session found but no env_dir. Using global session.")
        
        elif request.access_token:
            # Create a temporary directory for this token-based session
            temp_dir = tempfile.mkdtemp(prefix="az_session_")
            custom_env["AZURE_CONFIG_DIR"] = temp_dir
            
            print(f"🔐 Logging in with access token in isolated session: {temp_dir}")
            subprocess.run(
                ["az", "login", "--access-token", request.access_token],
                check=True,
                env=custom_env,
                capture_output=True
            )
        else:
            print("⚠️ No session ID or access token. Using global session (Not Isolated).")

        # Run assessment
        file_path = script.run_assessment(
            subscription_ids=[request.subscription_id] if request.subscription_id else None,
            env=custom_env
        )
        
        if file_path and os.path.exists(file_path):
            return {"status": "success", "file_path": os.path.abspath(file_path), "file_name": os.path.basename(file_path)}
        else:
            raise HTTPException(status_code=500, detail="Assessment script failed to generate file")

    except subprocess.CalledProcessError as e:
        raise HTTPException(status_code=401, detail=f"Authentication failed in isolated session: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Cleanup temporary directory
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)

@app.post("/api/analyze")
async def analyze_report(request: AIAssessmentRequest):
    """
    Endpoint to trigger AI analysis on an existing Excel report.
    Returns markdown analysis and conversation ID for chat.
    """
    if not os.path.exists(request.file_path):
        raise HTTPException(status_code=404, detail="Excel report not found")
        
    try:
        app.state.progress = {
            "step": "Extracting Context",
            "percentage": 10,
            "message": "Reading report sheets and preparing AI prompt"
        }

        print(f"[{datetime.now().isoformat()}] Starting AI analysis for: {request.file_path}")

        excel_file = pd.ExcelFile(request.file_path)
        all_data = {}
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(request.file_path, sheet_name=sheet_name)
            all_data[sheet_name] = df.to_csv(index=False)

        combined_summary = ""
        for sheet_name, csv_data in all_data.items():
            combined_summary += f"\n\n=== {sheet_name} ===\n{csv_data[:10000]}"
        if len(combined_summary) > 50000:
            combined_summary = combined_summary[:50000] + "\n...(truncated)"

        app.state.progress = {
            "step": "OpenAI Generation",
            "percentage": 40,
            "message": "Generating executive summary and migration recommendations"
        }

        prompt = f"""
        You are an Azure Migration Expert. Analyze the following comprehensive assessment data from multiple sheets:

        {combined_summary}

        Please provide a high-level executive summary in Markdown format, including:
        1. **Executive Summary**: Overall readiness, total resources, and key statistics.
        2. **Risk Analysis**: Identify high-risk resources and blockers.
        3. **Migration Strategy**: Recommended approach and priorities.
        4. **Next Steps**: Actionable items for the team.

        Keep it concise and professional.
        """

        client = openai.AzureOpenAI(
            azure_endpoint=AzureOpenAIConfig.ENDPOINT,
            api_key=AzureOpenAIConfig.API_KEY,
            api_version=AzureOpenAIConfig.API_VERSION
        )

        response = client.chat.completions.create(
            model=AzureOpenAIConfig.DEPLOYMENT_NAME,
            messages=[
                {"role": "system", "content": "You are a helpful Azure Migration Architect."},
                {"role": "user", "content": prompt}
            ]
        )

        analysis = response.choices[0].message.content
        conversation_id = str(uuid.uuid4())
        chat_sessions[conversation_id] = {
            "file_path": request.file_path,
            "data_summary": combined_summary,
            "history": [
                {"role": "system", "content": "You are a helpful Azure Migration Architect."},
                {"role": "user", "content": prompt},
                {"role": "assistant", "content": analysis}
            ]
        }

        app.state.progress = {"step": "Complete", "percentage": 100, "message": "Analysis finished"}

        return {
            "status": "success",
            "analysis": analysis,
            "conversation_id": conversation_id
        }
        
    except FileNotFoundError as e:
        print(f"Error: {str(e)}")
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        print(f"Error during AI analysis: {str(e)}")
        # Dump state for debugging
        app.state.progress = {"step": "Failed", "percentage": 0, "message": str(e)}
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate-architecture")
async def generate_architecture_endpoint(request: ArchitectureRequest):
    """
    Endpoint to generate an architecture diagram using Mermaid.js and Azure OpenAI.
    """
    if not os.path.exists(request.file_path):
        raise HTTPException(status_code=404, detail="Excel report not found")
        
    try:
        result = architecture_generator.generate_architecture_diagram(
            file_path=request.file_path,
            diagram_type=request.diagram_type,
            cloud_type=request.cloud_type
        )
        return {"status": "success", **result}
    except Exception as e:
        print(f"Error generating architecture: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to generate architecture: {str(e)}")


@app.post("/api/generate-architecture-upload")
async def generate_architecture_upload_endpoint(
    file: UploadFile = File(...),
    diagram_type: str = Form("network"),
    cloud_type: Optional[str] = Form(None),
):
    """
    Generate architecture directly from an uploaded Excel file.
    This endpoint is intended for diagram testing without running assessments.
    """
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported.")

    temp_file_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix="diagram_upload_") as tmp:
            content = await file.read()
            tmp.write(content)
            temp_file_path = tmp.name

        result = architecture_generator.generate_architecture_diagram(
            file_path=temp_file_path,
            diagram_type=diagram_type,
            cloud_type=cloud_type
        )
        return {
            "status": "success",
            "uploaded_file_name": file.filename,
            **result
        }
    except Exception as e:
        print(f"Error generating architecture from upload: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to generate architecture from upload: {str(e)}")
    finally:
        try:
            if temp_file_path and os.path.exists(temp_file_path):
                os.remove(temp_file_path)
        except Exception:
            pass

class DeviceCodeAssessmentRequest(BaseModel):
    session_id: str
    subscription_ids: list


@app.post("/api/assess-with-session")
async def run_assessment_with_session(request: DeviceCodeAssessmentRequest):
    """Run assessment using Device Code authenticated session."""
    from auth_api import auth_sessions
    
    session = auth_sessions.get(request.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found or expired")
    
    if session["status"] != "authenticated":
        raise HTTPException(status_code=401, detail="Session not authenticated")
    
    try:
        credential = session["credential"]
        file_path = script.run_assessment_with_credential(
            credential=credential,
            subscription_ids=request.subscription_ids
        )
        
        if file_path and os.path.exists(file_path):
            return {
                "status": "success",
                "file_path": os.path.abspath(file_path),
                "file_name": os.path.basename(file_path)
            }
        else:
            raise HTTPException(status_code=500, detail="Assessment failed to generate file")
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Assessment failed: {str(e)}")


class CliDeviceCodeAssessmentRequest(BaseModel):
    session_id: str
    subscription_ids: Optional[list] = None
    subscriptions: Optional[list] = None
    workbook_mode: str = "combined"


def update_job_progress(job_id: str, step: str, percentage: int, message: str = ""):
    """Helper function to update job progress"""
    if job_id in assessment_jobs:
        assessment_jobs[job_id]["progress"] = {
            "step": step,
            "percentage": percentage,
            "message": message
        }
        print(f"📊 Job {job_id}: {percentage}% - {step} {message}", flush=True)

def process_aws_assessment_job(job_id, request_dict):
    import time
    start_time = time.time()
    print(f"⚡ AWS Job Thread Started: {job_id}", flush=True)
    try:
        aws_access_key_id = request_dict.get("aws_access_key_id")
        aws_secret_access_key = request_dict.get("aws_secret_access_key")
        aws_session_token = request_dict.get("aws_session_token")
        regions = request_dict.get("region_names") or ([request_dict.get("region_name")] if request_dict.get("region_name") else [])
        
        if not regions:
             assessment_jobs[job_id]["status"] = "failed"
             assessment_jobs[job_id]["error"] = "No AWS regions specified."
             return

        all_resources = []
        total_regions = len(regions)
        
        from concurrent.futures import ThreadPoolExecutor
        
        # Function to process an individual region
        def process_region(region_info):
            idx, region = region_info
            region_base_pct = 10 + int((idx / total_regions) * 80)
            region_step_size = int(80 / total_regions)
            
            def aws_progress(step, pct, msg):
                # Map extractor-local percentage (0-100) to job-global percentage
                global_pct = region_base_pct + int((pct / 100) * region_step_size)
                update_job_progress(job_id, step, global_pct, f"[{region}] {msg}")

            try:
                update_job_progress(job_id, "Extracting", region_base_pct, f"Scanning region: {region} ({idx+1}/{total_regions})")
                
                extractor = aws_assessment.AwsResourceExtractor(
                    aws_access_key_id=aws_access_key_id,
                    aws_secret_access_key=aws_secret_access_key,
                    aws_session_token=aws_session_token,
                    region_name=region
                )
                return extractor.extract_resources(progress_callback=aws_progress)
            except Exception as e:
                print(f"❌ Error in region {region}: {e}")
                return []

        # Run extraction in parallel across regions
        with ThreadPoolExecutor(max_workers=min(len(regions), 8)) as executor:
            region_results = list(executor.map(process_region, enumerate(regions)))
            
        for resources in region_results:
            all_resources.extend(resources)
            
        if not all_resources:
            update_job_progress(job_id, "Completed", 100, "No resources found.")
            assessment_jobs[job_id]["status"] = "completed"
            assessment_jobs[job_id]["result"] = {
                "status": "warning",
                "message": "No resources found in the specified regions.",
                "file_path": None
            }
            return

        duration_sec = int(time.time() - start_time)
        duration_str = f"{duration_sec // 60}m {duration_sec % 60}s" if duration_sec > 60 else f"{duration_sec}s"
        
        update_job_progress(job_id, "Reporting", 90, "Generating Excel report...")
        file_path = aws_assessment.generate_aws_report(all_resources, duration=duration_str)
        
        update_job_progress(job_id, "Completed", 100, f"AWS Assessment complete! (Time: {duration_str})")
        assessment_jobs[job_id]["status"] = "completed"
        assessment_jobs[job_id]["duration"] = duration_str
        assessment_jobs[job_id]["result"] = {
            "status": "success",
            "file_path": os.path.abspath(file_path),
            "file_name": os.path.basename(file_path),
            "message": f"Successfully extracted {len(all_resources)} resources from {len(regions)} regions. (Total Time: {duration_str})",
            "duration": duration_str
        }
    except Exception as e:
        assessment_jobs[job_id]["status"] = "failed"
        assessment_jobs[job_id]["error"] = str(e)
        import traceback
        traceback.print_exc()

def process_gcp_assessment_job(job_id, request_dict):
    import time
    start_time = time.time()
    try:
        import gcp_assessment
        from google.oauth2.credentials import Credentials as UserCredentials
        import json as _json

        project_ids = request_dict.get("project_ids") or ([request_dict.get("project_id")] if request_dict.get("project_id") else [])
        gcp_session_id = request_dict.get("gcp_session_id")
        credentials_json = request_dict.get("credentials_json")

        extractor = None
        if gcp_session_id:
            session = gcp_oauth_sessions.get(gcp_session_id)
            if not session or session.get("status") != "completed":
                raise Exception("Invalid or expired GCP OAuth session")
            access_token = session.get("access_token")
            extractor = gcp_assessment.GcpResourceExtractor(
                project_ids=project_ids,
                credentials=UserCredentials(token=access_token)
            )
        elif credentials_json:
            credentials_dict = _json.loads(credentials_json)
            extractor = gcp_assessment.GcpResourceExtractor(
                project_ids=project_ids,
                credentials_json=credentials_dict
            )
        else:
             raise Exception("Missing GCP credentials")
        
        update_job_progress(job_id, "Extracting", 10, f"Found {len(project_ids)} projects to scan...")
        
        all_resources = []
        total_projects = len(project_ids)
        
        for idx, project_id in enumerate(project_ids):
            pct = 15 + int((idx / total_projects) * 65)
            update_job_progress(job_id, "Extracting", pct, f"Scanning project: {project_id} ({idx+1}/{total_projects})")
            
            # Use original extractor logic but scoped to one project at a time for progress reporting
            project_extractor = gcp_assessment.GcpResourceExtractor(
                project_ids=[project_id],
                credentials=extractor.credentials
            )
            resources = project_extractor.extract_resources()
            all_resources.extend(resources)
        
        if not all_resources:
             update_job_progress(job_id, "Completed", 100, "No resources found.")
             assessment_jobs[job_id]["status"] = "completed"
             assessment_jobs[job_id]["result"] = {
                "status": "warning",
                "message": "No GCP resources found. Ensure Cloud Asset API is enabled.",
                "file_path": None
             }
             return

        duration_sec = int(time.time() - start_time)
        duration_str = f"{duration_sec // 60}m {duration_sec % 60}s" if duration_sec > 60 else f"{duration_sec}s"

        update_job_progress(job_id, "Reporting", 80, "Generating Excel report...")
        file_path = gcp_assessment.generate_gcp_report(all_resources, duration=duration_str)
        
        if not file_path:
             assessment_jobs[job_id]["status"] = "failed"
             assessment_jobs[job_id]["error"] = "Failed to generate report file."
             return

        update_job_progress(job_id, "Completed", 100, f"GCP Assessment complete! (Time: {duration_str})")
        assessment_jobs[job_id]["status"] = "completed"
        assessment_jobs[job_id]["duration"] = duration_str
        assessment_jobs[job_id]["result"] = {
            "status": "success",
            "file_path": os.path.abspath(file_path),
            "file_name": os.path.basename(file_path),
            "message": f"Successfully extracted {len(all_resources)} GCP resources. (Total Time: {duration_str})",
            "resource_count": len(all_resources),
            "duration": duration_str
        }
    except Exception as e:
        assessment_jobs[job_id]["status"] = "failed"
        assessment_jobs[job_id]["error"] = str(e)
        import traceback
        traceback.print_exc()


def _subscription_report_name(subscription_name: str) -> str:
    """Create a safe, recognizable per-subscription workbook name."""
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", subscription_name or "subscription")
    safe_name = safe_name.strip("._-") or "subscription"
    return f"{ExcelExportConfig.FILE_NAME_PREFIX}_{safe_name}_{uuid.uuid4().hex[:8]}.xlsx"


PARALLEL_PROGRESS_EVENTS = (
    ("Fetching subscription names", 5, "Authenticating"),
    ("Querying Azure Resource Graph", 15, "Fetching resources"),
    ("Building NIC to Public IP map", 25, "Building network mappings"),
    ("Enriching VM resources", 30, "Enriching resources"),
    ("Exporting to Excel", 40, "Exporting workbook"),
    ("Formatting Excel file", 50, "Formatting workbook"),
    ("Adding migration plan with dependency analysis", 60, "Analyzing migration"),
    ("Appending quota info", 90, "Adding quota information"),
    ("Creating dashboard", 95, "Creating dashboard"),
)


def _update_parallel_worker_progress(job_id, subscription_id, subscription_name, percentage, step):
    """Update aggregate progress without allowing slower workers to move it backwards."""
    with assessment_jobs_lock:
        job = assessment_jobs.get(job_id)
        if not job or job.get("status") == "cancelled":
            return
        worker_progress = job.setdefault("_worker_progress", {})
        worker_progress[subscription_id] = max(
            percentage,
            worker_progress.get(subscription_id, 0),
        )
        total_workers = max(len(worker_progress), 1)
        aggregate = int(sum(worker_progress.values()) / total_workers)
        current_percentage = job.get("progress", {}).get("percentage", 0)
        job["progress"] = {
            "step": step,
            "percentage": min(max(aggregate, current_percentage), 99),
            "message": f"{subscription_name}: {step}",
        }


def _run_subscription_process(job_id, subscription_id, subscription_name, env_dir):
    """Run one subscription in an isolated Azure CLI environment."""
    worker_env_dir = tempfile.mkdtemp(prefix=f"az_assessment_{subscription_id[:8]}_")
    try:
        if env_dir and os.path.isdir(env_dir):
            shutil.copytree(env_dir, worker_env_dir, dirs_exist_ok=True)

        custom_env = os.environ.copy()
        custom_env["PYTHONUNBUFFERED"] = "1"
        custom_env["PYTHONIOENCODING"] = "utf-8"
        custom_env["AZURE_CONFIG_DIR"] = worker_env_dir

        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "script.py")
        process = subprocess.Popen(
            [
                sys.executable,
                script_path,
                "--subscription-ids",
                subscription_id,
                "--workbook-mode",
                "combined",
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            env=custom_env,
            encoding="utf-8",
            errors="replace",
        )
        assessment_jobs[job_id].setdefault("processes", {})[subscription_id] = process

        exported_paths = []
        for raw_line in iter(process.stdout.readline, ""):
            line = raw_line.strip()
            if line:
                print(f"[Script:{subscription_name}] {line}")
                for keyword, percentage, step in PARALLEL_PROGRESS_EVENTS:
                    if keyword in line:
                        _update_parallel_worker_progress(
                            job_id,
                            subscription_id,
                            subscription_name,
                            percentage,
                            step,
                        )
                        break
                if "Export complete:" in line:
                    exported_path = line.split("Export complete:", 1)[1].strip()
                    if exported_path not in exported_paths:
                        exported_paths.append(exported_path)

        process.wait()
        completed_paths = [path for path in exported_paths if os.path.isfile(path)]
        if process.returncode != 0 or not completed_paths:
            _update_parallel_worker_progress(
                job_id, subscription_id, subscription_name, 100, "Finished with errors"
            )
            return None, f"Assessment failed for {subscription_name} (exit code {process.returncode})"

        source_path = completed_paths[-1]
        named_path = os.path.join(
            os.path.dirname(source_path),
            _subscription_report_name(subscription_name),
        )
        os.replace(source_path, named_path)
        _update_parallel_worker_progress(
            job_id, subscription_id, subscription_name, 100, "Workbook ready"
        )
        return {
            "file_path": os.path.abspath(named_path),
            "file_name": os.path.basename(named_path),
            "subscription_id": subscription_id,
            "subscription_name": subscription_name,
        }, None
    except Exception as exc:
        _update_parallel_worker_progress(
            job_id, subscription_id, subscription_name, 100, "Finished with errors"
        )
        return None, f"Assessment failed for {subscription_name}: {exc}"
    finally:
        assessment_jobs.get(job_id, {}).get("processes", {}).pop(subscription_id, None)
        shutil.rmtree(worker_env_dir, ignore_errors=True)


def _process_parallel_subscription_jobs(job_id, subscription_ids, env_dir, subscriptions):
    """Generate subscription workbooks concurrently and publish each as it completes."""
    started_at = datetime.now()
    metadata = {
        str(item.get("subscription_id", "")).lower(): item.get("display_name")
        for item in (subscriptions or [])
        if isinstance(item, dict)
    }
    names = {
        subscription_id: metadata.get(subscription_id.lower()) or subscription_id
        for subscription_id in subscription_ids
    }
    result = {
        "status": "running",
        "message": "Generating subscription workbooks in parallel.",
        "workbook_mode": "per_subscription",
        "files": [],
        "total_files": len(subscription_ids),
        "completed_files": 0,
    }
    assessment_jobs[job_id]["result"] = result
    assessment_jobs[job_id]["processes"] = {}
    assessment_jobs[job_id]["_worker_progress"] = {
        subscription_id: 0 for subscription_id in subscription_ids
    }

    errors = []
    max_workers = min(len(subscription_ids), 4)
    update_job_progress(job_id, "Running in parallel", 5, f"Starting {len(subscription_ids)} subscription assessments...")

    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {
            pool.submit(
                _run_subscription_process,
                job_id,
                subscription_id,
                names[subscription_id],
                env_dir,
            ): subscription_id
            for subscription_id in subscription_ids
        }
        for future in as_completed(futures):
            file_result, error = future.result()
            if file_result:
                result["files"].append(file_result)
                result["completed_files"] = len(result["files"])
            if error:
                errors.append(error)

    assessment_jobs[job_id].pop("processes", None)
    assessment_jobs[job_id].pop("_worker_progress", None)
    if assessment_jobs[job_id].get("status") == "cancelled":
        return

    duration_sec = int((datetime.now() - started_at).total_seconds())
    duration_str = f"{duration_sec // 60}m {duration_sec % 60}s" if duration_sec > 60 else f"{duration_sec}s"
    result["duration"] = duration_str
    if errors:
        result["errors"] = errors

    if result["files"]:
        result["status"] = "warning" if errors else "success"
        result["message"] = (
            f"{len(result['files'])} of {len(subscription_ids)} workbooks completed."
            if errors
            else f"All {len(subscription_ids)} subscription workbooks completed."
        )
        assessment_jobs[job_id]["status"] = "completed"
        assessment_jobs[job_id]["duration"] = duration_str
        update_job_progress(job_id, "Completed", 100, result["message"])
    else:
        assessment_jobs[job_id]["status"] = "failed"
        assessment_jobs[job_id]["error"] = "; ".join(errors) or "No workbook was generated."


def process_assessment_job(
    job_id: str,
    subscription_ids: list,
    env_dir: str,
    workbook_mode: str = "combined",
    subscriptions: Optional[list] = None,
):
    """Run an assessment in cancellable subprocesses."""
    import time
    start_time = time.time()

    if workbook_mode == "per_subscription" and subscription_ids and len(subscription_ids) > 1:
        _process_parallel_subscription_jobs(
            job_id,
            subscription_ids,
            env_dir,
            subscriptions,
        )
        return
    
    print(f"🚀 Starting background job {job_id}")
    
    # Event-based progress mapping
    progress_events = {
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
    
    ai_analysis_started = False
    
    def monitor_output(text):
        nonlocal ai_analysis_started
        
        # Check for progress events
        for keyword, (percentage, step, message) in progress_events.items():
            if keyword in text:
                update_job_progress(job_id, step, percentage, message)
                if "Adding migration plan" in text:
                    ai_analysis_started = True
                return
        
        # Track AI analysis progress
        if ai_analysis_started and "📊 Progress:" in text:
            try:
                match = re.search(r'(\d+)/(\d+)\s+\((\d+)%\)', text)
                if match:
                    ai_completion_pct = int(match.group(3))
                    ui_progress = 45 + int((ai_completion_pct / 100) * 40)
                    ui_progress = min(ui_progress, 85)
                    
                    update_job_progress(
                        job_id, 
                        "Analyzing with AI", 
                        ui_progress,
                        f"Analyzing resources ({ai_completion_pct}% complete)..."
                    )
            except Exception:
                pass

    try:
        # Prepare env
        custom_env = os.environ.copy()
        custom_env["PYTHONUNBUFFERED"] = "1"
        custom_env["PYTHONIOENCODING"] = "utf-8"
        
        if env_dir:
            custom_env["AZURE_CONFIG_DIR"] = env_dir
            print(f"🔒 Using isolated session env for job {job_id}: {env_dir}")
        
        update_job_progress(job_id, "Initializing", 0, "Starting assessment...")
        
        # Locate script.py
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "script.py")
        
        cmd = [sys.executable, script_path]
        if subscription_ids:
            cmd.extend(["--subscription-ids"] + subscription_ids)
        cmd.extend(["--workbook-mode", workbook_mode])
            
        # Start subprocess
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            env=custom_env,
            encoding='utf-8',
            errors='replace'
        )
        
        # Store process in job
        if job_id in assessment_jobs:
            assessment_jobs[job_id]["process"] = process
            
        # Read output
        file_paths = []
        for line in iter(process.stdout.readline, ''):
            line = line.strip()
            if line:
                print(f"[Script] {line}")
                monitor_output(line)
                if "✅ Export complete:" in line:
                    parts = line.split(": ")
                    if len(parts) > 1:
                        exported_path = parts[1].strip()
                        if exported_path not in file_paths:
                            file_paths.append(exported_path)
                        
        process.wait()
        
        # Cleanup process ref
        if job_id in assessment_jobs and "process" in assessment_jobs[job_id]:
            del assessment_jobs[job_id]["process"]
            
        completed_files = [path for path in file_paths if os.path.exists(path)]
        if process.returncode == 0 and completed_files:
            duration_sec = int(time.time() - start_time)
            duration_str = f"{duration_sec // 60}m {duration_sec % 60}s" if duration_sec > 60 else f"{duration_sec}s"

            update_job_progress(job_id, "Completed", 100, f"Assessment complete! (Time: {duration_str})")
            assessment_jobs[job_id]["status"] = "completed"
            assessment_jobs[job_id]["duration"] = duration_str
            result = {
                "status": "success",
                "message": f"Azure assessment complete. (Total Time: {duration_str})",
                "duration": duration_str,
                "workbook_mode": workbook_mode,
                "files": [
                    {
                        "file_path": os.path.abspath(path),
                        "file_name": os.path.basename(path),
                        "subscription_id": (
                            subscription_ids[index]
                            if workbook_mode == "per_subscription" and subscription_ids and index < len(subscription_ids)
                            else None
                        ),
                    }
                    for index, path in enumerate(completed_files)
                ],
            }
            if workbook_mode == "combined" or len(completed_files) == 1:
                result["file_path"] = os.path.abspath(completed_files[0])
                result["file_name"] = os.path.basename(completed_files[0])
            assessment_jobs[job_id]["result"] = result
            print(f"✅ Job {job_id} completed successfully in {duration_str}")
        else:
            # Check if cancelled
            if assessment_jobs[job_id].get("status") == "cancelled":
                print(f"🚫 Job {job_id} was cancelled")
            else:
                assessment_jobs[job_id]["status"] = "failed"
                assessment_jobs[job_id]["error"] = "Assessment failed or no file generated"
                print(f"❌ Job {job_id} failed: Return code {process.returncode}")

    except Exception as e:
        assessment_jobs[job_id]["status"] = "failed"
        assessment_jobs[job_id]["error"] = str(e)
        print(f"❌ Job {job_id} failed with exception: {e}")
        import traceback
        traceback.print_exc()

@app.post("/api/cancel-assessment/{job_id}")
def cancel_assessment(job_id: str):
    job = assessment_jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    
    if "processes" in job:
        processes = list(job["processes"].values())
        for process in processes:
            try:
                process.terminate()
            except Exception:
                pass
        job["status"] = "cancelled"
        job["progress"]["message"] = "Assessment cancelled by user."
        return {"status": "success", "message": "Parallel assessments cancelled"}
    if "process" in job:
        process = job["process"]
        try:
            process.terminate()
            job["status"] = "cancelled"
            job["progress"]["message"] = "Assessment cancelled by user."
            return {"status": "success", "message": "Job cancelled"}
        except Exception as e:
             raise HTTPException(status_code=500, detail=f"Failed to cancel job: {e}")
    else:
        if job["status"] == "running":
             job["status"] = "cancelled"
             return {"status": "success", "message": "Job marked as cancelled"}
        return {"status": "warning", "message": "Job is not running or process not found"}

@app.get("/api/assessment-status/{job_id}")
def get_assessment_status(job_id: str):
    job = assessment_jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")
    
    # Return a copy without the process object which is not serializable
    job_response = job.copy()
    if "process" in job_response:
        del job_response["process"]
    if "processes" in job_response:
        del job_response["processes"]
    if "_worker_progress" in job_response:
        del job_response["_worker_progress"]
        
    return job_response


@app.post("/api/assess-with-cli-session")
def run_assessment_with_cli_session(request: CliDeviceCodeAssessmentRequest, background_tasks: BackgroundTasks):
    """
    Run assessment using CLI Device Code authenticated session.
    This uses the az CLI session that was authenticated via device code.
    """
    from auth_api import device_code_sessions
    
    session = device_code_sessions.get(request.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found or expired")
    
    if session["status"] != "completed":
        raise HTTPException(status_code=401, detail=f"Session not authenticated. Status: {session['status']}")

    if request.workbook_mode not in {"combined", "per_subscription"}:
        raise HTTPException(status_code=400, detail="workbook_mode must be 'combined' or 'per_subscription'")
    
    # Create a job ID
    job_id = str(uuid.uuid4())
    env_dir = session.get("env_dir")
    
    # Initialize job status
    assessment_jobs[job_id] = {
        "status": "running",
        "created_at": datetime.now().isoformat(),
        "session_id": request.session_id,
        "progress": {
            "step": "Queued",
            "percentage": 0,
            "message": "Assessment queued..."
        }
    }
    
    # Add to background tasks
    background_tasks.add_task(
        process_assessment_job, 
        job_id, 
        request.subscription_ids, 
        env_dir,
        request.workbook_mode,
        request.subscriptions
    )
    
    print(f"🚀 Started async assessment job {job_id} for session {request.session_id}")
    
    return {
        "status": "accepted", 
        "job_id": job_id, 
        "message": "Assessment started in background. Please poll /api/assessment-status/{job_id}"
    }


# Session storage for chat conversations (in-memory)
chat_sessions = {}

class ChatRequest(BaseModel):
    conversation_id: str
    message: str

@app.post("/api/chat")
async def chat(request: ChatRequest):
    if request.conversation_id not in chat_sessions:
        raise HTTPException(status_code=404, detail="Conversation not found")
    
    session = chat_sessions[request.conversation_id]
    
    try:
        # Add user message to history
        session["history"].append({"role": "user", "content": request.message})
        
        # Call OpenAI with full conversation history + data context
        client = openai.AzureOpenAI(
            azure_endpoint=AzureOpenAIConfig.ENDPOINT,
            api_key=AzureOpenAIConfig.API_KEY,
            api_version=AzureOpenAIConfig.API_VERSION
        )
        
        # Include data context in system message for better responses
        messages = [
            {"role": "system", "content": f"You are a helpful Azure Migration Architect. Here is the assessment data context:\n\n{session['data_summary'][:AIAnalysisConfig.MAX_DATA_CONTEXT_CHARS]}"},
        ] + session["history"][-AIAnalysisConfig.MAX_CHAT_HISTORY_MESSAGES:]  # Keep last N messages for context
        
        response = client.chat.completions.create(
            model=AzureOpenAIConfig.DEPLOYMENT_NAME,
            messages=messages
        )
        
        ai_response = response.choices[0].message.content
        
        # Add AI response to history
        session["history"].append({"role": "assistant", "content": ai_response})
        
        return {"status": "success", "response": ai_response}
    
    except Exception as e:
        print(f"Chat failed: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Chat failed: {str(e)}")

class AwsAssessmentRequest(BaseModel):
    aws_access_key_id: str
    aws_secret_access_key: str
    aws_session_token: Optional[str] = None
    region_name: Optional[str] = None # Backward compatibility
    region_names: Optional[list[str]] = None # New: Multi-region support

@app.post("/api/assess-aws")
async def run_aws_assessment(request: AwsAssessmentRequest, background_tasks: BackgroundTasks):
    try:
        job_id = str(uuid.uuid4())
        
        # Initialize job status
        assessment_jobs[job_id] = {
            "status": "running",
            "created_at": datetime.now().isoformat(),
            "progress": {
                "step": "Queued",
                "percentage": 0,
                "message": "AWS Assessment queued..."
            }
        }
        
        # We pass a dict because Pydantic models can't always be passed to background tasks easily if they contain session refs
        background_tasks.add_task(process_aws_assessment_job, job_id, request.dict())
        
        return {
            "status": "accepted",
            "job_id": job_id,
            "message": "AWS Assessment started in background."
        }
    except Exception as e:
        print(f"AWS Assessment startup failed: {e}")
        raise HTTPException(status_code=500, detail=f"AWS Assessment failed to start: {str(e)}")


class GcpAssessmentRequest(BaseModel):
    project_id: Optional[str] = None          # backward-compat — single project
    project_ids: Optional[list[str]] = None   # new — multiple projects
    credentials_json: Optional[str] = None
    gcp_session_id: Optional[str] = None

@app.post("/api/assess-gcp")
async def run_gcp_assessment(request: GcpAssessmentRequest, background_tasks: BackgroundTasks):
    """
    Run GCP to Azure migration assessment.
    Returns status: accepted and a job_id for polling.
    """
    try:
        job_id = str(uuid.uuid4())
        
        # Initialize job status
        assessment_jobs[job_id] = {
            "status": "running",
            "created_at": datetime.now().isoformat(),
            "progress": {
                "step": "Queued",
                "percentage": 0,
                "message": "GCP Assessment queued..."
            }
        }
        
        background_tasks.add_task(process_gcp_assessment_job, job_id, request.dict())
        
        return {
            "status": "accepted",
            "job_id": job_id,
            "message": "GCP Assessment started in background."
        }
    except Exception as e:
        print(f"GCP Assessment startup failed: {e}")
        raise HTTPException(status_code=500, detail=f"GCP Assessment failed to start: {str(e)}")
from fastapi.responses import FileResponse

@app.get("/api/download")
async def download_file(file_path: str):
    # Allow downloading only generated report files from known report directories.
    requested_path = os.path.abspath(file_path)
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    allowed_dirs = [
        os.path.abspath(os.path.join(project_root, "data")),
        os.path.abspath(tempfile.gettempdir()),
        os.path.abspath(os.path.join(project_root, "backend")),  # current GCP report location
    ]
    allowed_prefixes = (
        f"{ExcelExportConfig.FILE_NAME_PREFIX}_",
        "AWS_Assessment_",
        "gcp_to_azure_assessment_",
    )

    file_name = os.path.basename(requested_path)
    if not file_name.lower().endswith(".xlsx") or not file_name.startswith(allowed_prefixes):
        raise HTTPException(status_code=403, detail="Access denied")

    is_under_allowed_dir = any(
        os.path.commonpath([requested_path, allowed_dir]) == allowed_dir
        for allowed_dir in allowed_dirs
    )
    if not is_under_allowed_dir:
        raise HTTPException(status_code=403, detail="Access denied")

    if not os.path.exists(requested_path) or not os.path.isfile(requested_path):
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(
        requested_path,
        filename=file_name,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

app.add_middleware(
    CORSMiddleware,
    allow_origins=AppConfig.CORS_ALLOW_ORIGINS,
    allow_credentials=AppConfig.CORS_ALLOW_CREDENTIALS,
    allow_methods=AppConfig.CORS_ALLOW_METHODS,
    allow_headers=AppConfig.CORS_ALLOW_HEADERS,
)

# Mount static files
# We assume the frontend is built to ../frontend/dist
frontend_dist = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "frontend", "dist")

if os.path.exists(frontend_dist):
    app.mount("/assets", StaticFiles(directory=os.path.join(frontend_dist, "assets")), name="assets")

    @app.get("/{full_path:path}")
    async def serve_frontend(full_path: str):
        # API routes are already handled above because they are defined first
        # If it's a file in dist, serve it
        file_path = os.path.join(frontend_dist, full_path)
        if os.path.exists(file_path) and os.path.isfile(file_path):
            return FileResponse(file_path)
        
        # Otherwise serve index.html for SPA routing
        index_path = os.path.join(frontend_dist, "index.html")
        if os.path.exists(index_path):
            return FileResponse(index_path)
        
        return {"status": "error", "message": "Frontend not built"}
else:
    print("⚠️ Frontend dist directory not found. Running in API-only mode.")
