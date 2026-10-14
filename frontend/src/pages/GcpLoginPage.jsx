import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import axios from 'axios';

export default function GcpLoginPage() {
    const [formData, setFormData] = useState({
        project_id: '',
        credentials_json: '',
        use_file: false,
        credentials_file: null
    });
    const [authMode, setAuthMode] = useState('oauth');
    const [step, setStep] = useState('login'); // 'login' or 'select_project'
    const [projects, setProjects] = useState([]);
    const [loading, setLoading] = useState(false);
    const [oauthLoading, setOauthLoading] = useState(false);
    const [fetchingProjects, setFetchingProjects] = useState(false);
    const [sessionId, setSessionId] = useState(null);
    const [error, setError] = useState('');
    // Multi-project selection
    const [multiProjectMode, setMultiProjectMode] = useState(false);
    const [selectedProjects, setSelectedProjects] = useState([]);     // for multi-select
    const [singleSelectedProject, setSingleSelectedProject] = useState(null); // for single-select
    const [saProjectIds, setSaProjectIds] = useState(''); // Comma-separated projects for SA login
    const [multiProjectSAMode, setMultiProjectSAMode] = useState(false);
    const navigate = useNavigate();

    useEffect(() => {
        const onMessage = async (event) => {
            if (event.origin !== window.location.origin || !event.data?.type) {
                return;
            }

            if (event.data.type === 'gcp-oauth-success') {
                try {
                    const sid = event.data.sessionId;
                    setSessionId(sid);

                    const statusResp = await axios.get(`/api/auth/gcp/oauth-status/${sid}`);
                    if (statusResp.data?.status !== 'completed') {
                        throw new Error('Google sign-in not completed yet.');
                    }

                    // If we already have a project ID (from manual entry), go straight to assessment
                    if (formData.project_id.trim()) {
                        navigate('/assessment', {
                            state: {
                                gcpCredentials: {
                                    project_id: formData.project_id.trim(),
                                    gcp_session_id: sid
                                }
                            }
                        });
                        return;
                    }

                    // Otherwise, fetch projects
                    setStep('select_project');
                    setFetchingProjects(true);
                    setOauthLoading(false);

                    try {
                        const projectsResp = await axios.get(`/api/auth/gcp/projects/${sid}`);
                        setProjects(projectsResp.data);
                    } catch (err) {
                        console.error('Failed to fetch projects:', err);
                        setError('Authenticated successfully, but failed to fetch project list. Please enter Project ID manually.');
                        setStep('login');
                        setAuthMode('service_account'); // Fallback to manual entry
                    } finally {
                        setFetchingProjects(false);
                    }
                } catch (err) {
                    setError(err.response?.data?.detail || err.message || 'Failed to complete Google sign-in.');
                    setOauthLoading(false);
                }
            } else if (event.data.type === 'gcp-oauth-error') {
                setError(event.data.error || 'Google sign-in failed.');
                setOauthLoading(false);
            }
        };

        window.addEventListener('message', onMessage);
        return () => window.removeEventListener('message', onMessage);
    }, [formData.project_id, navigate]);

    const handleChange = (e) => {
        const { name, value, type, checked } = e.target;
        setFormData((prev) => ({
            ...prev,
            [name]: type === 'checkbox' ? checked : value
        }));
    };

    const handleFileChange = (e) => {
        const file = e.target.files[0];
        if (!file) {
            return;
        }
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                JSON.parse(event.target.result);
                setFormData((prev) => ({
                    ...prev,
                    credentials_file: file,
                    credentials_json: event.target.result
                }));
                setError('');
            } catch {
                setError('Invalid JSON file. Please upload a valid service account key.');
            }
        };
        reader.readAsText(file);
    };

    const handleGoogleSignIn = async () => {
        setError('');
        setOauthLoading(true);
        try {
            const response = await axios.post('/api/auth/gcp/oauth-start', {
                project_id: formData.project_id.trim() || null
            });
            const popup = window.open(
                response.data.auth_url,
                'gcp-oauth',
                'width=620,height=760,noopener,noreferrer'
            );
            if (!popup) {
                setOauthLoading(false);
                setError('Popup blocked. Please allow popups and try again.');
            }
        } catch (err) {
            setOauthLoading(false);
            setError(err.response?.data?.detail || 'Failed to start Google sign-in.');
        }
    };

    const handleServiceAccountLogin = async (e) => {
        e.preventDefault();
        setLoading(true);
        setError('');

        const projectList = multiProjectSAMode
            ? saProjectIds.split(',').map(p => p.trim()).filter(Boolean)
            : [formData.project_id.trim()];

        if (projectList.length === 0) {
            setError('At least one Project ID is required.');
            setLoading(false);
            return;
        }

        try {
            const payload = {
                project_ids: projectList,
                credentials_json: formData.credentials_json
            };

            const response = await axios.post('/api/auth/gcp/validate', payload);
            const normalizedProjectIds = response.data.project_ids || projectList;

            navigate('/assessment', {
                state: {
                    gcpCredentials: {
                        ...payload,
                        project_ids: normalizedProjectIds,
                        project_id: normalizedProjectIds[0] // backward compat
                    }
                }
            });
        } catch (err) {
            setError(err.response?.data?.detail || 'GCP authentication failed. Please check your credentials.');
        } finally {
            setLoading(false);
        }
    };

    if (step === 'select_project') {
        const handleProjectToggle = (projectId) => {
            setSelectedProjects(prev =>
                prev.includes(projectId)
                    ? prev.filter(id => id !== projectId)
                    : [...prev, projectId]
            );
        };

        const handleProceedWithProjects = () => {
            const ids = multiProjectMode ? selectedProjects : [singleSelectedProject];
            const validIds = ids.filter(Boolean);
            if (validIds.length === 0) return;
            navigate('/assessment', {
                state: {
                    gcpCredentials: {
                        project_ids: validIds,
                        gcp_session_id: sessionId
                    }
                }
            });
        };

        return (
            <div className="flex items-center justify-center min-h-screen bg-gray-900">
                <div className="w-full max-w-2xl p-8 space-y-6 bg-gray-800 rounded-xl shadow-2xl border border-gray-700">
                    <div className="text-center">
                        <h2 className="text-3xl font-bold text-white">Select GCP Project(s)</h2>
                        <p className="text-gray-400 mt-2">Choose the project(s) you want to assess for migration.</p>
                    </div>

                    {/* Multi-project toggle */}
                    <div className="flex items-center justify-between p-3 bg-gray-700/50 rounded-lg border border-gray-600">
                        <div>
                            <p className="text-sm font-medium text-white">Multi-Project Mode</p>
                            <p className="text-xs text-gray-400">Include multiple projects in one assessment sheet</p>
                        </div>
                        <button
                            type="button"
                            onClick={() => { setMultiProjectMode(m => !m); setSelectedProjects([]); setSingleSelectedProject(null); }}
                            className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors focus:outline-none ${multiProjectMode ? 'bg-blue-600' : 'bg-gray-500'}`}
                        >
                            <span className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${multiProjectMode ? 'translate-x-6' : 'translate-x-1'}`} />
                        </button>
                    </div>

                    {fetchingProjects ? (
                        <div className="py-12 flex flex-col items-center justify-center space-y-4">
                            <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin"></div>
                            <p className="text-gray-400">Fetching your GCP projects...</p>
                        </div>
                    ) : (
                        <div className="space-y-4">
                            {projects.length === 0 ? (
                                <div className="text-center py-8">
                                    <p className="text-gray-400">No active projects found in your account.</p>
                                    <button onClick={() => setStep('login')} className="mt-4 text-blue-400 hover:text-blue-300 underline">Back to login</button>
                                </div>
                            ) : (
                                <div className="max-h-96 overflow-y-auto pr-2 space-y-2 custom-scrollbar">
                                    {projects.map((p) => {
                                        const isSelected = multiProjectMode
                                            ? selectedProjects.includes(p.project_id)
                                            : singleSelectedProject === p.project_id;
                                        return (
                                            <button
                                                key={p.project_id}
                                                onClick={() => multiProjectMode ? handleProjectToggle(p.project_id) : setSingleSelectedProject(p.project_id)}
                                                className={`w-full text-left p-4 border-2 rounded-lg transition-all flex items-center justify-between group
                                                    ${isSelected
                                                        ? 'bg-blue-900/40 border-blue-500'
                                                        : 'bg-gray-700/50 border-gray-600 hover:border-blue-500 hover:bg-gray-700'
                                                    }`}
                                            >
                                                <div className="flex items-center gap-3">
                                                    {multiProjectMode ? (
                                                        <div className={`w-5 h-5 rounded border-2 flex items-center justify-center flex-shrink-0 ${isSelected ? 'bg-blue-500 border-blue-500' : 'border-gray-400'}`}>
                                                            {isSelected && <svg className="w-3 h-3 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M5 13l4 4L19 7" /></svg>}
                                                        </div>
                                                    ) : (
                                                        <div className={`w-5 h-5 rounded-full border-2 flex items-center justify-center flex-shrink-0 ${isSelected ? 'border-blue-500' : 'border-gray-400'}`}>
                                                            {isSelected && <div className="w-2.5 h-2.5 rounded-full bg-blue-500" />}
                                                        </div>
                                                    )}
                                                    <div>
                                                        <div className={`font-semibold ${isSelected ? 'text-blue-300' : 'text-white group-hover:text-blue-400'} transition`}>{p.display_name}</div>
                                                        <div className="text-xs text-gray-400">ID: {p.project_id}</div>
                                                    </div>
                                                </div>
                                                {isSelected && <span className="text-blue-400 text-sm font-medium">✓ Selected</span>}
                                            </button>
                                        );
                                    })}
                                </div>
                            )}

                            {/* Proceed button */}
                            {projects.length > 0 && (
                                <button
                                    onClick={handleProceedWithProjects}
                                    disabled={multiProjectMode ? selectedProjects.length === 0 : !singleSelectedProject}
                                    className={`w-full py-3 rounded-lg font-semibold transition-colors ${(multiProjectMode ? selectedProjects.length > 0 : !!singleSelectedProject)
                                        ? 'bg-blue-600 hover:bg-blue-500 text-white'
                                        : 'bg-gray-600 text-gray-400 cursor-not-allowed'
                                        }`}
                                >
                                    {multiProjectMode
                                        ? selectedProjects.length === 0
                                            ? 'Select at least one project'
                                            : `Run Assessment for ${selectedProjects.length} Project(s) →`
                                        : singleSelectedProject
                                            ? 'Run Assessment →'
                                            : 'Select a project'
                                    }
                                </button>
                            )}

                            <div className="pt-4 border-t border-gray-700 flex justify-between items-center text-sm">
                                <span className="text-gray-500">Authenticated as your Google account</span>
                                <button onClick={() => setStep('login')} className="text-gray-400 hover:text-white transition">Cancel</button>
                            </div>
                        </div>
                    )}
                </div>
            </div>
        );
    }

    return (
        <div className="flex items-center justify-center min-h-screen bg-gray-900">
            <div className="w-full max-w-2xl p-8 space-y-6 bg-gray-800 rounded-xl shadow-2xl border border-gray-700">
                <div className="text-center">
                    <h2 className="text-3xl font-bold text-white">GCP Migration Assessment</h2>
                    <p className="text-gray-400 mt-2">Sign in with Google or use service account credentials.</p>
                </div>

                {error && (
                    <div className="p-4 text-sm text-red-200 bg-red-900/50 border border-red-800 rounded-lg">
                        {error}
                    </div>
                )}

                <div className="flex space-x-2 bg-gray-700 p-1 rounded-lg">
                    <button
                        type="button"
                        onClick={() => setAuthMode('oauth')}
                        className={`flex-1 py-2 px-4 rounded-md text-sm font-medium transition ${authMode === 'oauth'
                            ? 'bg-blue-600 text-white'
                            : 'text-gray-300 hover:text-white'
                            }`}
                    >
                        Sign In (Recommended)
                    </button>
                    <button
                        type="button"
                        onClick={() => setAuthMode('service_account')}
                        className={`flex-1 py-2 px-4 rounded-md text-sm font-medium transition ${authMode === 'service_account'
                            ? 'bg-blue-600 text-white'
                            : 'text-gray-300 hover:text-white'
                            }`}
                    >
                        Service Account
                    </button>
                </div>

                {authMode === 'oauth' ? (
                    <div className="space-y-6 animate-in fade-in duration-300">
                        <div className="p-4 bg-blue-900/30 border border-blue-800 rounded-lg">
                            <h4 className="text-sm font-semibold text-blue-300 mb-2">Google Sign-In</h4>
                            <p className="text-xs text-gray-300 leading-relaxed">
                                You will be redirected to Google for authentication. After signing in, you will be able to select which GCP project to assess.
                            </p>
                        </div>

                        <div className="relative">
                            <div className="absolute inset-0 flex items-center" aria-hidden="true">
                                <div className="w-full border-t border-gray-700"></div>
                            </div>
                            <div className="relative flex justify-center text-xs uppercase">
                                <span className="bg-gray-800 px-2 text-gray-500 font-medium">Optional: Pre-select Project</span>
                            </div>
                        </div>

                        <div>
                            <label className="block mb-2 text-sm font-medium text-gray-300">
                                Project ID (if known)
                            </label>
                            <input
                                type="text"
                                name="project_id"
                                className="w-full px-4 py-3 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-white placeholder-gray-500 transition"
                                placeholder="my-gcp-project"
                                value={formData.project_id}
                                onChange={handleChange}
                            />
                        </div>

                        <button
                            type="button"
                            onClick={handleGoogleSignIn}
                            disabled={oauthLoading}
                            className={`w-full py-4 text-white font-bold rounded-lg shadow-lg transition duration-200 flex items-center justify-center space-x-3
                                ${oauthLoading ? 'bg-blue-700 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-500 hover:shadow-blue-500/30'}`}
                        >
                            {oauthLoading ? (
                                <>
                                    <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                                    <span>Waiting for Google sign-in...</span>
                                </>
                            ) : (
                                <>
                                    <svg className="w-5 h-5 fill-current" viewBox="0 0 24 24">
                                        <path d="M12.48 10.92v3.28h7.84c-.24 1.84-.908 3.152-1.908 4.152-1.228 1.228-3.14 2.532-7.84 2.532-7.072 0-12.72-5.744-12.72-12.816s5.648-12.816 12.72-12.816c3.824 0 6.64 1.504 8.71 3.52l2.304-2.304c-2.352-2.24-5.392-3.936-11.012-3.936-9.92 0-18.04 8.12-18.04 18.04s8.12 18.04 18.04 18.04c5.36 0 9.42-1.776 12.52-5.016 3.2-3.2 4.2-7.664 4.2-11.36 0-.856-.072-1.672-.2-2.48h-16.52z" />
                                    </svg>
                                    <span>Sign in with Google</span>
                                </>
                            )}
                        </button>
                    </div>
                ) : (
                    <form onSubmit={handleServiceAccountLogin} className="space-y-5 animate-in slide-in-from-bottom-4 duration-300">
                        <div className="space-y-4">
                            {/* Multi-project toggle for SA */}
                            <div className="flex items-center justify-between p-3 bg-gray-700/50 rounded-lg border border-gray-600 mb-4">
                                <div>
                                    <p className="text-sm font-medium text-white">Multi-Project Mode</p>
                                    <p className="text-xs text-gray-400">Include multiple projects in one assessment sheet</p>
                                </div>
                                <button
                                    type="button"
                                    onClick={() => setMultiProjectSAMode(!multiProjectSAMode)}
                                    className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors focus:outline-none ${multiProjectSAMode ? 'bg-blue-600' : 'bg-gray-500'}`}
                                >
                                    <span className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${multiProjectSAMode ? 'translate-x-6' : 'translate-x-1'}`} />
                                </button>
                            </div>

                            {multiProjectSAMode ? (
                                <div>
                                    <label className="block mb-2 text-sm font-medium text-gray-300">
                                        Project IDs (comma-separated) <span className="text-red-500">*</span>
                                    </label>
                                    <textarea
                                        required
                                        className="w-full px-4 py-3 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-white placeholder-gray-500 transition"
                                        placeholder="project-host-1, project-service-2"
                                        rows="2"
                                        value={saProjectIds}
                                        onChange={(e) => setSaProjectIds(e.target.value)}
                                    />
                                    <p className="mt-1 text-xs text-gray-500 italic">Example: host-project, service-project</p>
                                </div>
                            ) : (
                                <div>
                                    <label className="block mb-2 text-sm font-medium text-gray-300">
                                        Project ID <span className="text-red-500">*</span>
                                    </label>
                                    <input
                                        type="text"
                                        name="project_id"
                                        required
                                        className="w-full px-4 py-3 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-white placeholder-gray-500 transition"
                                        placeholder="my-gcp-project"
                                        value={formData.project_id}
                                        onChange={handleChange}
                                    />
                                </div>
                            )}

                            <label className="block text-sm font-medium text-gray-300">
                                Service Account Credentials <span className="text-red-500">*</span>
                            </label>

                            <div className="flex space-x-2 bg-gray-700 p-1 rounded-lg">
                                <button
                                    type="button"
                                    onClick={() => setFormData((prev) => ({ ...prev, use_file: false }))}
                                    className={`flex-1 py-2 px-4 rounded-md text-sm font-medium transition ${!formData.use_file
                                        ? 'bg-blue-600 text-white'
                                        : 'text-gray-300 hover:text-white'
                                        }`}
                                >
                                    Paste JSON
                                </button>
                                <button
                                    type="button"
                                    onClick={() => setFormData((prev) => ({ ...prev, use_file: true }))}
                                    className={`flex-1 py-2 px-4 rounded-md text-sm font-medium transition ${formData.use_file
                                        ? 'bg-blue-600 text-white'
                                        : 'text-gray-300 hover:text-white'
                                        }`}
                                >
                                    Upload File
                                </button>
                            </div>

                            {!formData.use_file ? (
                                <textarea
                                    name="credentials_json"
                                    required={!formData.use_file}
                                    rows="8"
                                    className="w-full px-4 py-3 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-white placeholder-gray-500 font-mono text-xs transition"
                                    placeholder='{"type": "service_account", "project_id": "...", ...}'
                                    value={formData.credentials_json}
                                    onChange={handleChange}
                                />
                            ) : (
                                <div className="flex items-center justify-center w-full">
                                    <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-gray-600 border-dashed rounded-lg cursor-pointer bg-gray-700 hover:bg-gray-600 transition">
                                        <div className="flex flex-col items-center justify-center pt-5 pb-6">
                                            <p className="mb-2 text-sm text-gray-400">
                                                <span className="font-semibold">Click to upload</span> or drag and drop
                                            </p>
                                            <p className="text-xs text-gray-500">Service Account JSON file</p>
                                            {formData.credentials_file && (
                                                <p className="mt-2 text-xs text-green-400">
                                                    {formData.credentials_file.name}
                                                </p>
                                            )}
                                        </div>
                                        <input
                                            type="file"
                                            accept=".json"
                                            onChange={handleFileChange}
                                            className="hidden"
                                            required={formData.use_file}
                                        />
                                    </label>
                                </div>
                            )}
                        </div>

                        <button
                            type="submit"
                            disabled={loading}
                            className={`w-full py-4 text-white font-bold rounded-lg shadow-lg transition duration-200 
                                ${loading ? 'bg-blue-700 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-500 hover:shadow-blue-500/30'}`}
                        >
                            {loading ? 'Verifying...' : 'Connect Service Account'}
                        </button>
                    </form>
                )}
            </div>
        </div>
    );
}
