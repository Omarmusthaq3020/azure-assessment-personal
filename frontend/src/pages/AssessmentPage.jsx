import { useState, useEffect, useRef } from 'react';
import { useLocation, useNavigate } from 'react-router-dom';
import axios from 'axios';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import ProgressIndicator from '../components/ProgressIndicator';
import MermaidDiagram from '../components/MermaidDiagram';

export default function AssessmentPage() {
    const [formData, setFormData] = useState({
        subscription_id: '',
        migration_type: 'azure_to_azure'
    });
    const [loading, setLoading] = useState(false);
    const [result, setResult] = useState(null);
    const [error, setError] = useState('');
    const [analyzing, setAnalyzing] = useState(false);
    const [analysis, setAnalysis] = useState('');
    const [conversationId, setConversationId] = useState('');
    const [chatHistory, setChatHistory] = useState([]);
    const [chatMessage, setChatMessage] = useState('');
    const [chatLoading, setChatLoading] = useState(false);

    // Architecture Diagram State
    const [diagramType, setDiagramType] = useState('network');
    const [diagramCode, setDiagramCode] = useState('');
    const [drawioXml, setDrawioXml] = useState('');
    const [generatingDiagram, setGeneratingDiagram] = useState(false);

    // Device Code session data
    const location = useLocation();
    const navigate = useNavigate();
    const [sessionId, setSessionId] = useState(null);
    const [subscriptionIds, setSubscriptionIds] = useState([]);
    const [subscriptions, setSubscriptions] = useState([]);
    const [authMethod, setAuthMethod] = useState(null); // 'sdk' or 'az-cli'
    const [progress, setProgress] = useState(null); // Progress tracking
    const [elapsedTime, setElapsedTime] = useState(0);
    const timerRef = useRef(null);

    const hasAutoTriggered = useRef(false);
    const jobIdRef = useRef(null);
    const loadingRef = useRef(false);

    useEffect(() => {
        loadingRef.current = loading;
    }, [loading]);

    useEffect(() => {
        const handleUnload = () => {
            if (jobIdRef.current && loadingRef.current) {
                const url = `/api/cancel-assessment/${jobIdRef.current}`;
                navigator.sendBeacon(url);
            }
        };

        window.addEventListener('beforeunload', handleUnload);
        return () => {
            window.removeEventListener('beforeunload', handleUnload);
        };
    }, []);

    useEffect(() => {
        // Only run auto-config once per mount
        if (hasAutoTriggered.current) return;

        let detectedType = null;

        // 1. Detect Azure Device Code session
        if (location.state?.sessionId) {
            console.log('🔗 Azure session detected');
            setSessionId(location.state.sessionId);
            setSubscriptionIds(location.state?.subscriptionIds || []);
            setSubscriptions(location.state?.subscriptions || []);
            setAuthMethod(location.state?.authMethod || 'sdk');
            detectedType = 'azure_to_azure';
        }

        // 2. Detect AWS credentials
        if (location.state?.awsCredentials) {
            console.log('🔗 AWS credentials detected');
            detectedType = 'aws_to_azure';
        }

        // 3. Detect GCP credentials
        if (location.state?.gcpCredentials) {
            console.log('🔗 GCP credentials detected');
            detectedType = 'gcp_to_azure';
        }

        if (detectedType) {
            setFormData(prev => ({ ...prev, migration_type: detectedType }));
            hasAutoTriggered.current = true;
        }
    }, [location.state]);

    const handleChange = (e) => {
        setFormData({ ...formData, [e.target.name]: e.target.value });
    };

    const handleAnalyze = async () => {
        if (!result?.file_path) return;
        setAnalyzing(true);
        setAnalysis('');
        setChatHistory([]);
        setConversationId('');
        try {
            const response = await axios.post('/api/analyze', {
                file_path: result.file_path
            });
            setAnalysis(response.data.analysis);
            setConversationId(response.data.conversation_id);
        } catch (err) {
            console.error(err);
            setError('AI Analysis failed: ' + (err.response?.data?.detail || err.message));
        } finally {
            setAnalyzing(false);
        }
    };

    const handleGenerateArchitecture = async () => {
        if (!result?.file_path) return;
        setGeneratingDiagram(true);
        setDiagramCode('');
        setDrawioXml('');
        setError('');
        try {
            const response = await axios.post('/api/generate-architecture', {
                file_path: result.file_path,
                diagram_type: diagramType,
                cloud_type: formData.migration_type.split('_')[0]
            });
            setDiagramCode(response.data.diagram);
            setDrawioXml(response.data.drawio_xml || '');
        } catch (err) {
            console.error(err);
            setError('Failed to generate architecture diagram: ' + (err.response?.data?.detail || err.message));
        } finally {
            setGeneratingDiagram(false);
        }
    };

    const handleChat = async () => {
        if (!chatMessage.trim() || !conversationId) return;

        const userMessage = chatMessage;
        setChatMessage('');
        setChatHistory(prev => [...prev, { role: 'user', content: userMessage }]);
        setChatLoading(true);

        try {
            const response = await axios.post('/api/chat', {
                conversation_id: conversationId,
                message: userMessage
            });
            setChatHistory(prev => [...prev, { role: 'assistant', content: response.data.response }]);
        } catch (err) {
            console.error(err);
            setError('Chat failed: ' + (err.response?.data?.detail || err.message));
        } finally {
            setChatLoading(false);
        }
    };

    const cancelAssessment = async (jobId) => {
        try {
            await axios.post(`/api/cancel-assessment/${jobId}`);
            setLoading(false);
            setProgress(null);
            if (timerRef.current) clearInterval(timerRef.current);
            setError('Assessment cancelled by user.');
        } catch (err) {
            console.error('Failed to cancel:', err);
        }
    };

    const startTimer = () => {
        setElapsedTime(0);
        if (timerRef.current) clearInterval(timerRef.current);
        timerRef.current = setInterval(() => {
            setElapsedTime(prev => prev + 1);
        }, 1000);
    };

    const stopTimer = () => {
        if (timerRef.current) {
            clearInterval(timerRef.current);
            timerRef.current = null;
        }
    };

    const formatTime = (seconds) => {
        const mins = Math.floor(seconds / 60);
        const secs = seconds % 60;
        return `${mins}:${secs.toString().padStart(2, '0')}`;
    };

    const pollAssessmentStatus = async (jobId) => {
        let consecutiveErrors = 0;
        const maxConsecutiveErrors = 6;
        const pollInterval = setInterval(async () => {
            try {
                const response = await axios.get(`/api/assessment-status/${jobId}`);
                const job = response.data;
                consecutiveErrors = 0;

                if (job.status === 'completed') {
                    clearInterval(pollInterval);
                    stopTimer();
                    jobIdRef.current = null;
                    setResult(job.result);
                    setLoading(false);
                    setProgress(null);
                } else if (job.status === 'failed') {
                    clearInterval(pollInterval);
                    stopTimer();
                    jobIdRef.current = null;
                    setError(job.error || 'Assessment failed');
                    setLoading(false);
                    setProgress(null);
                } else if (job.status === 'running' && job.progress) {
                    // Update progress while running
                    setProgress(job.progress);
                }
                // If running, continue polling
            } catch (err) {
                console.error('Polling error:', err);
                const statusCode = err?.response?.status;
                const isTransient = !statusCode || statusCode === 404 || statusCode >= 500;
                if (isTransient && consecutiveErrors < maxConsecutiveErrors) {
                    consecutiveErrors += 1;
                    return;
                }
                clearInterval(pollInterval);
                stopTimer();
                jobIdRef.current = null;
                setError('Failed to check assessment status');
                setLoading(false);
            }
        }, 5000); // Poll every 5 seconds for responsive progress updates
    };

    const handleRunAssessment = async (overrideType = null) => {
        const currentType = overrideType || formData.migration_type;
        console.log('--- Submitting Assessment Request ---');
        console.log('Type:', currentType);
        
        setLoading(true);
        setError('');
        setResult(null);
        startTimer();

        try {
            // CASE A: AWS to Azure
            if (currentType === 'aws_to_azure') {
                const awsCreds = location.state?.awsCredentials;
                if (!awsCreds) {
                    throw new Error('AWS credentials missing. Please re-login.');
                }
                const response = await axios.post('/api/assess-aws', awsCreds);
                if (response.data.status === 'accepted') {
                    jobIdRef.current = response.data.job_id;
                    pollAssessmentStatus(response.data.job_id);
                    return; 
                } else if (response.data.status === 'success' || response.data.status === 'warning') {
                    setResult(response.data);
                } else {
                    setError(response.data.message || 'AWS Assessment failed.');
                }
            } 
            
            // CASE B: GCP to Azure
            else if (currentType === 'gcp_to_azure') {
                const gcpCreds = location.state?.gcpCredentials;
                if (!gcpCreds) {
                    throw new Error('GCP credentials missing. Please re-login.');
                }
                const payload = {
                    gcp_session_id: gcpCreds.gcp_session_id,
                    credentials_json: gcpCreds.credentials_json,
                    project_ids: gcpCreds.project_ids || (gcpCreds.project_id ? [gcpCreds.project_id] : undefined),
                    project_id: gcpCreds.project_id
                };
                const response = await axios.post('/api/assess-gcp', payload);
                if (response.data.status === 'accepted') {
                    jobIdRef.current = response.data.job_id;
                    pollAssessmentStatus(response.data.job_id);
                    return;
                } else if (response.data.status === 'success' || response.data.status === 'warning') {
                    setResult(response.data);
                } else {
                    setError(response.data.message || 'GCP Assessment failed.');
                }
            }

            // CASE C: Azure (with Session)
            else if (sessionId) {
                let response;
                if (authMethod === 'az-cli') {
                    response = await axios.post('/api/assess-with-cli-session', {
                        session_id: sessionId,
                        subscription_ids: subscriptionIds.length > 0 ? subscriptionIds : null
                    });
                    if (response.data.status === 'accepted') {
                        jobIdRef.current = response.data.job_id;
                        pollAssessmentStatus(response.data.job_id);
                        return; // Async polling handles the rest
                    }
                } else {
                    response = await axios.post('/api/assess-with-session', {
                        session_id: sessionId,
                        subscription_ids: subscriptionIds
                    });
                }

                if (response?.data.status === 'success') {
                    setResult(response.data);
                } else {
                    setError(response?.data.message || 'Azure Assessment failed.');
                }
            }

            // CASE D: Legacy Azure (Token fallback)
            else if (currentType === 'azure_to_azure') {
                const token = localStorage.getItem('azure_access_token');
                const response = await axios.post('/api/assess', {
                    ...formData,
                    access_token: token
                });
                if (response.data.status === 'success' || response.data.status === 'warning') {
                    setResult(response.data);
                } else {
                    setError(response.data.message || 'Assessment failed.');
                }
            }

            // UNHANDLED
            else {
                setError(`No credentials found for ${currentType}. Please log in again.`);
            }

        } catch (err) {
            console.error('Assessment execution crashed:', err);
            setError(err.response?.data?.detail || err.message || 'Failed to run assessment.');
            setLoading(false);
        }
        // finally block removed as it was prematurely setting loading:false for background jobs
    };

    const handleDownload = () => {
        if (result?.file_path) {
            window.open(`/api/download?file_path=${encodeURIComponent(result.file_path)}`, '_blank');
        }
    };

    return (
        <div className="min-h-screen bg-gray-900 text-gray-100 p-8">
            <div className="max-w-4xl mx-auto">
                <header className="mb-10 text-center">
                    <h1 className="text-4xl font-bold text-white mb-2">Assessment Dashboard</h1>
                    <p className="text-gray-400">Configure and run your migration assessment</p>
                    <div className="mt-4">
                        <button
                            onClick={() => navigate('/diagram-playground')}
                            className="px-4 py-2 bg-emerald-700 hover:bg-emerald-600 text-white rounded-lg border border-emerald-600 text-sm"
                        >
                            Open Diagram Playground
                        </button>
                    </div>
                </header>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                    {/* Configuration Panel */}
                    <div className="bg-gray-800 p-6 rounded-xl shadow-lg border border-gray-700">
                        <h2 className="text-xl font-semibold mb-4 text-blue-400">Configuration</h2>
                        <div className="space-y-4">
                            {/* Show selected subscriptions if available */}
                            {subscriptionIds && subscriptionIds.length > 0 && (
                                <div>
                                    <div className="flex justify-between items-center mb-2">
                                        <label className="block text-sm font-medium text-gray-300">
                                            Selected Subscriptions
                                        </label>
                                        <button
                                            onClick={() => navigate('/azure-login', {
                                                state: {
                                                    sessionId: sessionId,
                                                    returnToSubscriptions: true,
                                                    subscriptionIds: subscriptionIds
                                                }
                                            })}
                                            className="text-xs text-blue-400 hover:text-blue-300 underline"
                                        >
                                            Change
                                        </button>
                                    </div>
                                    <div className="bg-gray-700 border border-gray-600 rounded-lg p-3">
                                        <p className="text-sm text-green-400 mb-2">
                                            ✓ {subscriptionIds.length} subscription{subscriptionIds.length > 1 ? 's' : ''} selected
                                        </p>
                                        <div className="max-h-32 overflow-y-auto space-y-2">
                                            {subscriptions.length > 0 ? (
                                                subscriptions.map((sub, index) => (
                                                    <div key={index} className="text-sm">
                                                        <p className="text-white font-medium">{sub.display_name}</p>
                                                        <p className="text-xs text-gray-400 font-mono">{sub.subscription_id}</p>
                                                    </div>
                                                ))
                                            ) : (
                                                subscriptionIds.map((id, index) => (
                                                    <p key={index} className="text-xs text-gray-400 font-mono">
                                                        {id}
                                                    </p>
                                                ))
                                            )}
                                        </div>
                                    </div>
                                </div>
                            )}

                            <div>
                                <label className="block mb-2 text-sm font-medium text-gray-300">Migration Type</label>
                                <select
                                    name="migration_type"
                                    className="w-full px-4 py-2 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 text-white"
                                    value={formData.migration_type}
                                    onChange={handleChange}
                                >
                                    <option value="azure_to_azure">Azure to Azure</option>
                                    <option value="aws_to_azure">AWS to Azure</option>
                                    <option value="gcp_to_azure">GCP to Azure</option>
                                </select>
                            </div>

                            <div className="flex gap-4 mt-4">
                                <button
                                    onClick={() => handleRunAssessment()}
                                    disabled={loading}
                                    className={`flex-1 py-3 px-6 rounded-lg font-bold text-white transition-all shadow-lg ${
                                        loading 
                                        ? 'bg-gray-600 cursor-not-allowed opacity-70' 
                                        : 'bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 hover:scale-[1.02] active:scale-95'
                                    }`}
                                >
                                    {loading ? (
                                        <div className="flex items-center justify-center gap-2">
                                            <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
                                            Running...
                                        </div>
                                    ) : 'Run Assessment'}
                                </button>

                                {loading && jobIdRef.current && (
                                    <button
                                        onClick={() => cancelAssessment(jobIdRef.current)}
                                        className="px-4 py-3 bg-red-900/30 border border-red-800 text-red-400 font-bold rounded-lg hover:bg-red-900/50 transition-all flex items-center gap-2"
                                        title="Cancel Assessment"
                                    >
                                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" />
                                        </svg>
                                        Cancel
                                    </button>
                                )}
                            </div>
                        </div>
                    </div >

                    {/* Status Panel */}
                    < div className="bg-gray-800 p-6 rounded-xl shadow-lg border border-gray-700 flex flex-col justify-center items-center text-center" >
                        {!loading && !result && !error && (
                            <div className="text-gray-500">
                                <svg className="w-16 h-16 mx-auto mb-4 opacity-50" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                                </svg>
                                <p>Ready to generate report</p>
                            </div>
                        )
                        }

                        {
                            loading && (
                                <>
                                    <div className="flex justify-between items-center mb-6">
                                        <h2 className="text-2xl font-bold text-white flex items-center">
                                            <svg className="w-8 h-8 mr-3 text-blue-400 animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                                            </svg>
                                            Running Assessment...
                                        </h2>
                                        <div className="px-4 py-2 bg-gray-800 rounded-lg border border-gray-700">
                                            <span className="text-gray-400 text-sm font-medium mr-2">Time Elapsed:</span>
                                            <span className="text-blue-400 font-mono text-lg font-bold">{formatTime(elapsedTime)}</span>
                                        </div>
                                    </div>
                                    <ProgressIndicator progress={progress} />
                                </>
                            )
                        }

                        {
                            error && (
                                <div className="text-red-400">
                                    <svg className="w-16 h-16 mx-auto mb-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                                    </svg>
                                    <p className="font-semibold">Error</p>
                                    <p className="text-sm mt-2">{error}</p>
                                </div>
                            )
                        }

                        {
                            result && (
                                <div className="space-y-4 w-full">
                                    <div className="w-16 h-16 bg-green-900/50 rounded-full flex items-center justify-center mx-auto border border-green-500/50">
                                        <svg className="w-8 h-8 text-green-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" />
                                        </svg>
                                    </div>
                                    <h3 className="text-xl font-bold text-white">Assessment Complete!</h3>
                                    {result.duration && (
                                        <p className="text-gray-400 text-sm mt-1 font-medium animate-fade-in">
                                            Completed in: <span className="text-blue-400 font-bold">{result.duration}</span>
                                        </p>
                                    )}
                                    {result.status === 'warning' && (
                                        <p className="text-yellow-400 text-sm">{result.message}</p>
                                    )}
                                    <div className="bg-gray-900 p-4 rounded-lg text-left text-sm font-mono text-gray-300 break-all">
                                        {result.file_name}
                                    </div>
                                    <button
                                        onClick={handleDownload}
                                        className="w-full py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg font-medium transition shadow-lg hover:shadow-blue-500/30 flex items-center justify-center mb-3"
                                    >
                                        <svg className="w-5 h-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                                        </svg>
                                        Download Report
                                    </button>

                                    <button
                                        onClick={handleAnalyze}
                                        disabled={analyzing}
                                        className={`w-full py-2 text-white rounded-lg font-medium transition shadow-lg flex items-center justify-center ${analyzing ? 'bg-purple-800 cursor-not-allowed' : 'bg-purple-600 hover:bg-purple-500 hover:shadow-purple-500/30'}`}
                                    >
                                        {analyzing ? (
                                            <>
                                                <svg className="w-5 h-5 mr-2 animate-spin" viewBox="0 0 24 24">
                                                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                                                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                                                </svg>
                                                Analyzing with AI...
                                            </>
                                        ) : (
                                            <>
                                                <svg className="w-5 h-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 10V3L4 14h7v7l9-11h-7z" />
                                                </svg>
                                                Analyze with AI
                                            </>
                                        )}
                                    </button>

                                    <div className="pt-3 border-t border-gray-700 mt-4 flex flex-col gap-2">
                                        <div className="flex gap-2">
                                            <select 
                                                value={diagramType} 
                                                onChange={(e) => setDiagramType(e.target.value)}
                                                className="bg-gray-700 border border-gray-600 rounded-lg text-white text-sm px-2 py-2 flex-1 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                                            >
                                                <option value="network">Network Topology</option>
                                                <option value="resource_groups">Resource Groups</option>
                                                <option value="migration">Migration Readiness</option>
                                            </select>
                                            
                                            <button
                                                onClick={handleGenerateArchitecture}
                                                disabled={generatingDiagram}
                                                className={`flex-1 py-2 text-white rounded-lg text-sm font-medium transition shadow-lg flex items-center justify-center ${generatingDiagram ? 'bg-emerald-800 cursor-not-allowed' : 'bg-emerald-600 hover:bg-emerald-500 hover:shadow-emerald-500/30'}`}
                                            >
                                                {generatingDiagram ? (
                                                    <>
                                                        <svg className="w-4 h-4 mr-2 animate-spin" viewBox="0 0 24 24">
                                                            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                                                            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                                                        </svg>
                                                        Generating...
                                                    </>
                                                ) : (
                                                    <>🏗️ Generate Diagram</>
                                                )}
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            )
                        }
                    </div >
                </div >

                {/* Architecture Diagram Result */}
                {diagramCode && (
                    <div className="mt-8 bg-gray-800 p-8 rounded-xl shadow-lg border border-gray-700 animate-fade-in max-w-full overflow-hidden">
                        <div className="flex items-center justify-between mb-6">
                            <div className="flex items-center">
                                <div className="p-2 bg-emerald-900/50 rounded-lg mr-4">
                                    <svg className="w-8 h-8 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 5a1 1 0 011-1h14a1 1 0 011 1v2a1 1 0 01-1 1H5a1 1 0 01-1-1V5zM4 13a1 1 0 011-1h6a1 1 0 011 1v6a1 1 0 01-1 1H5a1 1 0 01-1-1v-6zM16 13a1 1 0 011-1h2a1 1 0 011 1v6a1 1 0 01-1 1h-2a1 1 0 01-1-1v-6z" />
                                    </svg>
                                </div>
                                <h2 className="text-2xl font-bold text-white">
                                    {diagramType === 'network' ? 'Network Topology' : 
                                     diagramType === 'resource_groups' ? 'Resource Groups Overview' : 
                                     'Migration Readiness'}
                                </h2>
                            </div>
                        </div>
                        <MermaidDiagram code={diagramCode} drawioXml={drawioXml} />
                    </div>
                )}

                {/* AI Analysis Result */}
                {
                    analysis && (
                        <div className="mt-8 bg-gray-800 p-8 rounded-xl shadow-lg border border-gray-700 animate-fade-in max-w-full overflow-hidden">
                            <div className="flex items-center mb-6">
                                <div className="p-2 bg-purple-900/50 rounded-lg mr-4">
                                    <svg className="w-8 h-8 text-purple-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19.428 15.428a2 2 0 00-1.022-.547l-2.384-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z" />
                                    </svg>
                                </div>
                                <h2 className="text-2xl font-bold text-white">AI Executive Summary</h2>
                            </div>

                            <div className="prose prose-invert max-w-none break-words overflow-wrap-anywhere">
                                <ReactMarkdown
                                    remarkPlugins={[remarkGfm]}
                                    components={{
                                        h1: ({ node, ...props }) => <h1 className="text-white text-2xl font-semibold mt-6 mb-4 pb-2 border-b border-gray-700 break-words" {...props} />,
                                        h2: ({ node, ...props }) => <h2 className="text-white text-xl font-semibold mt-8 mb-3 break-words" {...props} />,
                                        h3: ({ node, ...props }) => <h3 className="text-white text-lg font-semibold mt-6 mb-2 break-words" {...props} />,
                                        p: ({ node, ...props }) => <p className="text-gray-300 my-4 leading-7 break-words" {...props} />,
                                        strong: ({ node, ...props }) => <strong className="text-white font-semibold break-words" {...props} />,
                                        ul: ({ node, ...props }) => <ul className="text-gray-300 my-4 space-y-2 pl-6 list-disc break-words" {...props} />,
                                        ol: ({ node, ...props }) => <ol className="text-gray-300 my-4 space-y-2 pl-6 list-decimal break-words" {...props} />,
                                        li: ({ node, ...props }) => <li className="text-gray-300 leading-7 break-words" {...props} />,
                                        table: ({ node, ...props }) => (
                                            <div className="overflow-x-auto my-6">
                                                <table className="min-w-full border border-gray-700 rounded-lg" {...props} />
                                            </div>
                                        ),
                                        thead: ({ node, ...props }) => <thead className="bg-gray-900/50" {...props} />,
                                        tbody: ({ node, ...props }) => <tbody className="divide-y divide-gray-700" {...props} />,
                                        tr: ({ node, ...props }) => <tr className="hover:bg-gray-800/30" {...props} />,
                                        th: ({ node, ...props }) => <th className="px-4 py-3 text-left text-sm font-semibold text-white border-b border-gray-700 break-words" {...props} />,
                                        td: ({ node, ...props }) => <td className="px-4 py-3 text-sm text-gray-300 break-words" {...props} />,
                                    }}
                                >
                                    {analysis}
                                </ReactMarkdown>
                            </div>

                            {/* Chat Interface */}
                            <div className="mt-8 border-t border-gray-700 pt-6">
                                <h3 className="text-xl font-bold text-white mb-4 flex items-center">
                                    <svg className="w-6 h-6 mr-2 text-purple-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M8 10h.01M12 10h.01M16 10h.01M9 16H5a2 2 0 01-2-2V6a2 2 0 012-2h14a2 2 0 012 2v8a2 2 0 01-2 2h-5l-5 5v-5z" />
                                    </svg>
                                    Ask Follow-up Questions
                                </h3>

                                {/* Chat History */}
                                {chatHistory.length > 0 && (
                                    <div className="mb-4 max-h-96 overflow-y-auto space-y-3 bg-gray-900/50 rounded-lg p-4">
                                        {chatHistory.map((msg, idx) => (
                                            <div key={idx} className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'}`}>
                                                <div className={`max-w-3xl rounded-lg px-4 py-2 ${msg.role === 'user' ? 'bg-blue-600 text-white' : 'bg-gray-700 text-gray-200'}`}>
                                                    {msg.role === 'user' ? (
                                                        <p className="text-sm leading-relaxed">{msg.content}</p>
                                                    ) : (
                                                        <ReactMarkdown
                                                            remarkPlugins={[remarkGfm]}
                                                            components={{
                                                                h1: ({ node, ...props }) => <h1 className="text-white text-base font-semibold mt-3 mb-2" {...props} />,
                                                                h2: ({ node, ...props }) => <h2 className="text-white text-sm font-semibold mt-3 mb-2" {...props} />,
                                                                h3: ({ node, ...props }) => <h3 className="text-white text-sm font-semibold mt-2 mb-1" {...props} />,
                                                                p: ({ node, ...props }) => <p className="text-gray-200 my-2 leading-6" {...props} />,
                                                                strong: ({ node, ...props }) => <strong className="text-white font-medium" {...props} />,
                                                                ul: ({ node, ...props }) => <ul className="text-gray-200 my-2 space-y-1 pl-5 list-disc" {...props} />,
                                                                ol: ({ node, ...props }) => <ol className="text-gray-200 my-2 space-y-1 pl-5 list-decimal" {...props} />,
                                                                li: ({ node, ...props }) => <li className="text-gray-200 leading-6" {...props} />
                                                            }}
                                                        >
                                                            {msg.content}
                                                        </ReactMarkdown>
                                                    )}
                                                </div>
                                            </div>
                                        ))}
                                        {chatLoading && (
                                            <div className="flex justify-start">
                                                <div className="bg-gray-700 text-gray-200 rounded-lg px-4 py-2">
                                                    <div className="flex items-center space-x-2">
                                                        <div className="w-2 h-2 bg-gray-400 rounded-full animate-bounce"></div>
                                                        <div className="w-2 h-2 bg-gray-400 rounded-full animate-bounce" style={{ animationDelay: '0.2s' }}></div>
                                                        <div className="w-2 h-2 bg-gray-400 rounded-full animate-bounce" style={{ animationDelay: '0.4s' }}></div>
                                                    </div>
                                                </div>
                                            </div>
                                        )}
                                    </div>
                                )}

                                {/* Chat Input */}
                                <div className="flex gap-2">
                                    <input
                                        type="text"
                                        value={chatMessage}
                                        onChange={(e) => setChatMessage(e.target.value)}
                                        onKeyPress={(e) => e.key === 'Enter' && handleChat()}
                                        placeholder="Ask a question about your assessment..."
                                        className="flex-1 px-4 py-2 bg-gray-700 text-white rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                                        disabled={chatLoading || !conversationId}
                                    />
                                    <button
                                        onClick={handleChat}
                                        disabled={chatLoading || !chatMessage.trim() || !conversationId}
                                        className="px-6 py-2 bg-purple-600 hover:bg-purple-500 text-white rounded-lg font-medium transition disabled:opacity-50 disabled:cursor-not-allowed flex items-center"
                                    >
                                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8" />
                                        </svg>
                                    </button>
                                </div>
                            </div>
                        </div>
                    )
                }
            </div >
        </div >
    );
}
