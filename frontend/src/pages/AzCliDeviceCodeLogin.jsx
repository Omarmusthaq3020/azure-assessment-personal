import { useState, useEffect } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';

const AZURE_SESSION_STORAGE_KEY = 'azure_cli_assessment_session';

export default function AzCliDeviceCodeLogin() {
    const [step, setStep] = useState('start'); // 'start', 'device-code', 'loading-subscriptions', 'select-subscriptions'
    const [sessionId, setSessionId] = useState('');
    const [deviceCodeInfo, setDeviceCodeInfo] = useState(null);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState('');
    const [copied, setCopied] = useState(false);
    const [tenantId, setTenantId] = useState('');
    const [subscriptions, setSubscriptions] = useState([]);
    const [selectedSubscriptions, setSelectedSubscriptions] = useState([]);
    const [workbookMode, setWorkbookMode] = useState('combined');

    const navigate = useNavigate();
    const location = useLocation();

    useEffect(() => {
        if (location.state?.returnToSubscriptions && location.state?.sessionId) {
            console.log('Returning to subscription selection with session:', location.state.sessionId);
            setSessionId(location.state.sessionId);
            if (location.state.subscriptionIds) {
                setSelectedSubscriptions(location.state.subscriptionIds);
            }
            const tid = location.state.tenantId;
            if (tid) setTenantId(tid);
            if (location.state.workbookMode) {
                setWorkbookMode(location.state.workbookMode);
            }
            fetchSubscriptions(location.state.sessionId, tid);
            return;
        }

        const savedSession = sessionStorage.getItem(AZURE_SESSION_STORAGE_KEY);
        if (!savedSession) return;

        try {
            const saved = JSON.parse(savedSession);
            if (!saved.sessionId) return;

            setSessionId(saved.sessionId);
            setTenantId(saved.tenantId || '');
            setSelectedSubscriptions(saved.subscriptionIds || []);
            setWorkbookMode(saved.workbookMode || 'combined');
            if (saved.authenticated && saved.subscriptions?.length) {
                setSubscriptions(saved.subscriptions);
                setStep('select-subscriptions');
                return;
            }

            setStep('loading-subscriptions');
            fetch(`/api/auth/device-code-status/${saved.sessionId}`)
                .then((response) => response.ok ? response.json() : Promise.reject(new Error('Session expired')))
                .then((status) => {
                    if (status.status === 'completed') {
                        fetchSubscriptions(saved.sessionId, saved.tenantId || null);
                    } else if (status.status === 'pending' && saved.deviceCodeInfo) {
                        setDeviceCodeInfo(saved.deviceCodeInfo);
                        setStep('device-code');
                        pollAuthStatus(saved.sessionId, saved.tenantId || null);
                    } else {
                        sessionStorage.removeItem(AZURE_SESSION_STORAGE_KEY);
                        setStep('start');
                    }
                })
                .catch(() => {
                    sessionStorage.removeItem(AZURE_SESSION_STORAGE_KEY);
                    setStep('start');
                });
        } catch {
            sessionStorage.removeItem(AZURE_SESSION_STORAGE_KEY);
        }
        // The polling functions intentionally own their timer lifecycle.
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [location]);

    const handleStartAuth = async () => {
        const tenantIdToAuth = tenantId.trim() || null;
        setLoading(true);
        setError('');
        setDeviceCodeInfo(null);

        try {
            console.log('📡 Calling /api/auth/device-code-start...');

            // Build URL with tenant_id if provided
            let url = '/api/auth/device-code-start';
            // Ensure tenantIdToAuth is a string and not an event object
            if (tenantIdToAuth && typeof tenantIdToAuth === 'string') {
                url += `?tenant_id=${encodeURIComponent(tenantIdToAuth)}`;
            }

            const response = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }
            });

            console.log('📥 Response status:', response.status);

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                console.error('❌ Error response:', errorData);
                throw new Error(errorData.detail || 'Failed to start az CLI authentication');
            }

            const data = await response.json();
            console.log('✅ Received data:', data);

            setSessionId(data.session_id);
            const receivedCodeInfo = {
                userCode: data.user_code,
                verificationUrl: data.verification_url,
                message: data.message,
                authenticatingTenant: tenantIdToAuth // Track which tenant we're authenticating to
            };
            setDeviceCodeInfo(receivedCodeInfo);
            setStep('device-code');
            sessionStorage.setItem(AZURE_SESSION_STORAGE_KEY, JSON.stringify({
                sessionId: data.session_id,
                tenantId: tenantIdToAuth || '',
                deviceCodeInfo: receivedCodeInfo
            }));

            console.log('📋 Device Code:', data.user_code);
            console.log('🔗 Verification URL:', data.verification_url);
            // Start polling for authentication status
            pollAuthStatus(data.session_id, tenantIdToAuth);

        } catch (err) {
            console.error('❌ Error:', err);
            setError(err.message);
            setStep('start');
        } finally {
            setLoading(false);
        }
    };

    const pollAuthStatus = async (sid, authenticatingTenantId = null) => {
        let completed = false;
        let inFlight = false;

        const cleanup = (pollInterval, timeoutId, focusHandler, visibilityHandler) => {
            clearInterval(pollInterval);
            clearTimeout(timeoutId);
            window.removeEventListener('focus', focusHandler);
            document.removeEventListener('visibilitychange', visibilityHandler);
        };

        const checkStatus = async (pollInterval, timeoutId, focusHandler, visibilityHandler) => {
            if (completed || inFlight) return;
            inFlight = true;

            try {
                const response = await fetch(`/api/auth/device-code-status/${sid}`);
                const data = await response.json();

                if (data.status === 'completed') {
                    completed = true;
                    cleanup(pollInterval, timeoutId, focusHandler, visibilityHandler);
                    setStep('loading-subscriptions');
                    const saved = JSON.parse(sessionStorage.getItem(AZURE_SESSION_STORAGE_KEY) || '{}');
                    sessionStorage.setItem(AZURE_SESSION_STORAGE_KEY, JSON.stringify({
                        ...saved,
                        sessionId: sid,
                        tenantId: authenticatingTenantId || '',
                        authenticated: true
                    }));
                    console.log(
                        authenticatingTenantId
                            ? `Authenticated to tenant ${authenticatingTenantId}, fetching subscriptions...`
                            : 'Authenticated to the default tenant, fetching subscriptions...'
                    );
                    await fetchSubscriptions(sid, authenticatingTenantId);
                } else if (data.status === 'failed') {
                    completed = true;
                    cleanup(pollInterval, timeoutId, focusHandler, visibilityHandler);
                    setError(data.error || 'Authentication failed');
                    setStep('start');
                }
            } catch (err) {
                console.error('Polling error:', err);
            } finally {
                inFlight = false;
            }
        };

        const focusHandler = () => {
            setStep('loading-subscriptions');
            checkStatus(pollInterval, timeoutId, focusHandler, visibilityHandler);
        };

        const visibilityHandler = () => {
            if (document.visibilityState === 'visible') {
                setStep('loading-subscriptions');
                checkStatus(pollInterval, timeoutId, focusHandler, visibilityHandler);
            }
        };

        const pollInterval = setInterval(() => {
            checkStatus(pollInterval, timeoutId, focusHandler, visibilityHandler);
        }, 300);

        window.addEventListener('focus', focusHandler);
        document.addEventListener('visibilitychange', visibilityHandler);

        const timeoutId = setTimeout(
            () => cleanup(pollInterval, timeoutId, focusHandler, visibilityHandler),
            900000
        );

        await checkStatus(pollInterval, timeoutId, focusHandler, visibilityHandler);
    };
    const fetchSubscriptions = async (sid, tenantId = null) => {
        setLoading(true);
        setError('');
        try {
            console.log('📋 Fetching subscriptions...');
            let url = `/api/auth/device-code-subscriptions/${sid}`;
            if (tenantId) {
                url += `?tenant_id=${tenantId}`;
            }

            const response = await fetch(url);

            if (!response.ok) {
                const errText = await response.text();
                throw new Error(`Failed to fetch subscriptions (${response.status}): ${errText}`);
            }

            const text = await response.text();
            try {
                const subs = JSON.parse(text);
                setSubscriptions(subs);
                setStep('select-subscriptions');
                const saved = JSON.parse(sessionStorage.getItem(AZURE_SESSION_STORAGE_KEY) || '{}');
                sessionStorage.setItem(AZURE_SESSION_STORAGE_KEY, JSON.stringify({
                    ...saved,
                    sessionId: sid,
                    tenantId: tenantId || '',
                    subscriptions: subs,
                    authenticated: true
                }));
            } catch (e) {
                console.error('❌ JSON Parse Error:', e);
                console.error('Response text was:', text);
                throw new Error(`Invalid JSON response: ${e.message}`);
            }

        } catch (err) {
            console.error('❌ Error fetching subscriptions:', err);
            setError(`Error: ${err.message}`);
            // Don't restart immediately, let user see error
            // setStep('start');
        } finally {
            setLoading(false);
        }
    };

    const handleSubscriptionToggle = (subId) => {
        setSelectedSubscriptions(prev => {
            if (prev.includes(subId)) {
                return prev.filter(id => id !== subId);
            } else {
                return [...prev, subId];
            }
        });
    };

    const handleProceedToAssessment = () => {
        // Get full subscription details for selected IDs
        const selectedSubDetails = subscriptions.filter(sub =>
            selectedSubscriptions.includes(sub.subscription_id)
        );

        const sessionState = {
            authMethod: 'az-cli',
            sessionId,
            subscriptionIds: selectedSubscriptions,
            subscriptions: selectedSubDetails,
            tenantId: tenantId.trim(),
            workbookMode: selectedSubscriptions.length > 1 ? workbookMode : 'combined'
        };
        sessionStorage.setItem(AZURE_SESSION_STORAGE_KEY, JSON.stringify({
            ...sessionState,
            subscriptions,
            authenticated: true
        }));

        navigate('/assessment', {
            state: sessionState
        });
    };

    const copyToClipboard = (text) => {
        navigator.clipboard.writeText(text);
        setCopied(true);
        setTimeout(() => setCopied(false), 2000);
    };

    return (
        <div className="min-h-screen bg-gradient-to-br from-indigo-50 to-purple-100 flex items-center justify-center p-4">
            <div className="max-w-2xl w-full">

                {/* Start Step */}
                {step === 'start' && (
                    <div className="bg-white rounded-2xl shadow-xl p-8">
                        <div className="text-center mb-8">
                            <div className="inline-block p-4 bg-indigo-100 rounded-full mb-4">
                                <svg className="w-12 h-12 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 9l3 3-3 3m5 0h3M5 20h14a2 2 0 002-2V6a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
                                </svg>
                            </div>
                            <h1 className="text-3xl font-bold text-gray-900 mb-2">Azure Migration Assessment</h1>
                            <p className="text-gray-600">Sign in to start your assessment</p>
                        </div>

                        {error && (
                            <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700">
                                {error}
                            </div>
                        )}

                        <div className="space-y-6">
                            <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
                                <h3 className="font-semibold text-blue-900 mb-2">How it works</h3>
                                <p className="text-sm text-blue-800">
                                    Enter a tenant ID to authenticate directly to that tenant, or leave it blank to use
                                    your Azure account's default tenant.
                                </p>
                            </div>

                            <div>
                                <label htmlFor="tenant-id" className="block text-sm font-medium text-gray-700 mb-2">
                                    Azure Tenant ID <span className="font-normal text-gray-500">(optional)</span>
                                </label>
                                <input
                                    id="tenant-id"
                                    type="text"
                                    value={tenantId}
                                    onChange={(event) => setTenantId(event.target.value)}
                                    onKeyDown={(event) => {
                                        if (event.key === 'Enter') handleStartAuth();
                                    }}
                                    placeholder="00000000-0000-0000-0000-000000000000"
                                    autoComplete="off"
                                    className="w-full px-4 py-3 border border-gray-300 rounded-lg font-mono text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                                />
                                <p className="mt-2 text-xs text-gray-500">
                                    Leave blank for the default tenant. Find a specific ID in Microsoft Entra ID under Overview → Tenant ID.
                                </p>
                            </div>

                            <button
                                onClick={handleStartAuth}
                                disabled={loading}
                                className="w-full bg-indigo-600 text-white py-3 px-6 rounded-lg font-semibold hover:bg-indigo-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                            >
                                {loading
                                    ? 'Requesting device code...'
                                    : tenantId.trim()
                                        ? 'Sign in to this tenant'
                                        : 'Sign in with default tenant'}
                            </button>
                        </div>
                    </div>
                )}

                {/* Device Code Display Step */}
                {step === 'device-code' && deviceCodeInfo && (
                    <div className="bg-white rounded-2xl shadow-xl p-8">
                        <div className="text-center mb-8">
                            <div className="inline-block p-4 bg-green-100 rounded-full mb-4">
                                <svg className="w-12 h-12 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 18h.01M8 21h8a2 2 0 002-2V5a2 2 0 00-2-2H8a2 2 0 00-2 2v14a2 2 0 002 2z" />
                                </svg>
                            </div>
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">Complete Authentication</h2>
                            <p className="text-gray-600">Follow these steps to sign in</p>

                            {deviceCodeInfo.authenticatingTenant && (
                                <div className="mt-4 inline-block px-4 py-2 bg-blue-100 border border-blue-300 rounded-lg">
                                    <p className="text-sm text-blue-800">
                                        🎯 Authenticating to specific tenant
                                    </p>
                                    <p className="text-xs text-blue-600 font-mono mt-1">
                                        {deviceCodeInfo.authenticatingTenant}
                                    </p>
                                </div>
                            )}
                        </div>

                        <div className="space-y-6">
                            <div className="bg-blue-50 border border-blue-200 rounded-lg p-6">
                                <p className="text-sm font-medium text-gray-700 mb-4">Step 1: Open the login page</p>
                                <div className="flex items-center gap-3">
                                    <input
                                        type="text"
                                        value={deviceCodeInfo.verificationUrl}
                                        readOnly
                                        className="flex-1 px-4 py-2 bg-white border border-gray-300 rounded-lg text-sm"
                                    />
                                    <button
                                        onClick={() => window.open(deviceCodeInfo.verificationUrl, '_blank')}
                                        className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
                                    >
                                        Open
                                    </button>
                                </div>
                            </div>

                            <div className="bg-green-50 border border-green-200 rounded-lg p-6">
                                <p className="text-sm font-medium text-gray-700 mb-4">Step 2: Enter this code</p>
                                <div className="flex items-center gap-3">
                                    <div className="flex-1 px-6 py-4 bg-white border-2 border-green-300 rounded-lg">
                                        <p className="text-3xl font-mono font-bold text-center text-gray-900 tracking-wider">
                                            {deviceCodeInfo.userCode}
                                        </p>
                                    </div>
                                    <button
                                        onClick={() => copyToClipboard(deviceCodeInfo.userCode)}
                                        className={`px-4 py-2 rounded-lg transition-colors ${copied
                                            ? 'bg-green-700 text-white'
                                            : 'bg-green-600 text-white hover:bg-green-700'
                                            }`}
                                    >
                                        {copied ? 'Copied!' : 'Copy'}
                                    </button>
                                </div>
                            </div>

                            <div className="text-center">
                                <div className="inline-flex items-center gap-2 text-gray-600">
                                    <div className="animate-spin h-5 w-5 border-2 border-indigo-600 border-t-transparent rounded-full"></div>
                                    <span>Waiting for you to complete authentication...</span>
                                </div>
                            </div>
                        </div>
                    </div>
                )}

                {/* Post-auth transition step */}
                {step === 'loading-subscriptions' && (
                    <div className="bg-white rounded-2xl shadow-xl p-8">
                        <div className="text-center">
                            <div className="inline-block p-4 bg-green-100 rounded-full mb-4">
                                <svg className="w-12 h-12 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                                </svg>
                            </div>
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">Authentication Successful</h2>
                            <p className="text-gray-600 mb-6">
                                {tenantId
                                    ? `Loading subscriptions for tenant ${tenantId}...`
                                    : 'Loading subscriptions for your default tenant...'}
                            </p>
                            <div className="inline-flex items-center gap-2 text-gray-600">
                                <div className="animate-spin h-5 w-5 border-2 border-indigo-600 border-t-transparent rounded-full"></div>
                                <span>Please wait</span>
                            </div>
                        </div>
                    </div>
                )}

                {/* Subscription Selection Step */}
                {step === 'select-subscriptions' && (
                    <div className="bg-white rounded-2xl shadow-xl p-8">
                        <div className="text-center mb-8">
                            <div className="inline-block p-4 bg-purple-100 rounded-full mb-4">
                                <svg className="w-12 h-12 text-purple-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01" />
                                </svg>
                            </div>
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">Select Subscriptions</h2>
                            <p className="text-gray-600">Choose which subscriptions to assess</p>

                            <p className="mt-2 text-xs text-gray-500">
                                {tenantId ? `Tenant: ${tenantId}` : 'Default tenant'}
                            </p>
                        </div>

                        <div className="space-y-4">
                            <div className="flex items-center justify-between mb-4 pb-4 border-b border-gray-200">
                                <p className="text-sm text-gray-600">
                                    {selectedSubscriptions.length} of {subscriptions.length} selected
                                </p>
                                <button
                                    onClick={() => {
                                        if (selectedSubscriptions.length === subscriptions.length) {
                                            setSelectedSubscriptions([]);
                                        } else {
                                            setSelectedSubscriptions(subscriptions.map(s => s.subscription_id));
                                        }
                                    }}
                                    className="text-sm text-indigo-600 hover:text-indigo-700 font-medium"
                                >
                                    {selectedSubscriptions.length === subscriptions.length ? 'Deselect All' : 'Select All'}
                                </button>
                            </div>

                            <div className="max-h-96 overflow-y-auto space-y-2">
                                {subscriptions.map((sub) => (
                                    <label
                                        key={sub.subscription_id}
                                        className={`flex items-start p-4 rounded-lg border-2 cursor-pointer transition-all ${selectedSubscriptions.includes(sub.subscription_id)
                                            ? 'border-indigo-500 bg-indigo-50'
                                            : 'border-gray-200 hover:border-gray-300'
                                            }`}
                                    >
                                        <input
                                            type="checkbox"
                                            checked={selectedSubscriptions.includes(sub.subscription_id)}
                                            onChange={() => handleSubscriptionToggle(sub.subscription_id)}
                                            className="mt-1 h-4 w-4 text-indigo-600 rounded focus:ring-indigo-500"
                                        />
                                        <div className="ml-3 flex-1">
                                            <p className="font-medium text-gray-900">{sub.display_name}</p>
                                            <p className="text-sm text-gray-500 font-mono">{sub.subscription_id}</p>
                                            <div className="flex items-center gap-2 mt-1">
                                                <span className={`text-xs px-2 py-0.5 rounded ${sub.state === 'Enabled'
                                                    ? 'bg-green-100 text-green-700'
                                                    : 'bg-gray-100 text-gray-700'
                                                    }`}>
                                                    {sub.state}
                                                </span>
                                                {sub.is_default && (
                                                    <span className="text-xs px-2 py-0.5 rounded bg-blue-100 text-blue-700">
                                                        Default
                                                    </span>
                                                )}
                                                {/* Show tenant ID if debug needed, or just keep clean */}
                                            </div>
                                        </div>
                                    </label>
                                ))}
                            </div>

                            {selectedSubscriptions.length > 1 && (
                                <fieldset className="mt-6 border border-gray-200 rounded-xl p-4">
                                    <legend className="px-2 text-sm font-semibold text-gray-800">Workbook output</legend>
                                    <div className="space-y-3">
                                        <label className="flex items-start gap-3 cursor-pointer">
                                            <input
                                                type="radio"
                                                name="workbook-mode"
                                                value="combined"
                                                checked={workbookMode === 'combined'}
                                                onChange={(event) => setWorkbookMode(event.target.value)}
                                                className="mt-1 h-4 w-4 text-indigo-600"
                                            />
                                            <span>
                                                <span className="block font-medium text-gray-900">One combined workbook</span>
                                                <span className="block text-xs text-gray-500">All selected subscriptions in one report.</span>
                                            </span>
                                        </label>
                                        <label className="flex items-start gap-3 cursor-pointer">
                                            <input
                                                type="radio"
                                                name="workbook-mode"
                                                value="per_subscription"
                                                checked={workbookMode === 'per_subscription'}
                                                onChange={(event) => setWorkbookMode(event.target.value)}
                                                className="mt-1 h-4 w-4 text-indigo-600"
                                            />
                                            <span>
                                                <span className="block font-medium text-gray-900">One workbook per subscription</span>
                                                <span className="block text-xs text-gray-500">Generate a separate downloadable report for each subscription.</span>
                                            </span>
                                        </label>
                                    </div>
                                </fieldset>
                            )}

                            <button
                                onClick={handleProceedToAssessment}
                                disabled={selectedSubscriptions.length === 0}
                                className="w-full mt-6 bg-indigo-600 text-white py-3 px-6 rounded-lg font-semibold hover:bg-indigo-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                            >
                                {selectedSubscriptions.length === 0
                                    ? 'Select at least one subscription'
                                    : `Proceed with ${selectedSubscriptions.length} subscription${selectedSubscriptions.length > 1 ? 's' : ''}`
                                }
                            </button>
                        </div>
                    </div>
                )}

            </div>
        </div >
    );
}


