import { useState, useEffect } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';

export default function AzCliDeviceCodeLogin() {
    const [step, setStep] = useState('start'); // 'start', 'device-code', 'loading-tenants', 'select-tenant', 'select-subscriptions', 'completed'
    const [sessionId, setSessionId] = useState('');
    const [deviceCodeInfo, setDeviceCodeInfo] = useState(null);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState('');
    const [copied, setCopied] = useState(false);
    const [tenants, setTenants] = useState([]);
    const [selectedTenant, setSelectedTenant] = useState('');
    const [subscriptions, setSubscriptions] = useState([]);
    const [selectedSubscriptions, setSelectedSubscriptions] = useState([]);

    const navigate = useNavigate();
    const location = useLocation();

    useEffect(() => {
        if (location.state?.returnToSubscriptions && location.state?.sessionId) {
            console.log('Returning to subscription selection with session:', location.state.sessionId);
            setSessionId(location.state.sessionId);
            if (location.state.subscriptionIds) {
                setSelectedSubscriptions(location.state.subscriptionIds);
            }
            // If we have a stored tenant ID, use it, otherwise just fetch default
            const tid = location.state.tenantId;
            if (tid) setSelectedTenant(tid);
            fetchSubscriptions(location.state.sessionId, tid);
        }
    }, [location]);

    const handleStartAuth = async (tenantIdToAuth = null) => {
        console.log('🚀 Starting CLI authentication...', tenantIdToAuth ? `for tenant ${tenantIdToAuth}` : 'for home tenant');
        setLoading(true);
        setError('');
        setStep('device-code');
        setDeviceCodeInfo({
            userCode: 'Loading...',
            verificationUrl: 'https://microsoft.com/devicelogin',
            message: 'Starting authentication. Fetching device code...',
            authenticatingTenant: tenantIdToAuth
        });

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
            setDeviceCodeInfo({
                userCode: data.user_code,
                verificationUrl: data.verification_url,
                message: data.message,
                authenticatingTenant: tenantIdToAuth // Track which tenant we're authenticating to
            });

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
                    setStep('loading-tenants');

                    if (authenticatingTenantId) {
                        console.log(`Authenticated to tenant ${authenticatingTenantId}, fetching subscriptions...`);
                        setSelectedTenant(authenticatingTenantId);
                        await fetchSubscriptions(sid, authenticatingTenantId);
                    } else {
                        await fetchTenants(sid);
                    }
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
            checkStatus(pollInterval, timeoutId, focusHandler, visibilityHandler);
        };

        const visibilityHandler = () => {
            if (document.visibilityState === 'visible') {
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
    const fetchTenants = async (sid) => {
        try {
            console.log('📋 Fetching tenants...');
            const response = await fetch(`/api/auth/device-code-tenants/${sid}`);

            if (!response.ok) {
                // If endpoint fails (e.g. old backend), fallback to fetching subscriptions directly
                console.warn('⚠️ Fetch tenants failed, falling back to direct subscription fetch');
                await fetchSubscriptions(sid);
                return;
            }

            const data = await response.json();
            console.log('✅ Tenants:', data);
            setTenants(data);

            if (data.length === 0) {
                // No tenants explicitly found, fallback
                await fetchSubscriptions(sid);
            } else if (data.length === 1 && data[0].subscription_count > 0) {
                // Only one tenant and it has subscriptions, auto-select
                const tid = data[0].tenant_id;
                setSelectedTenant(tid);
                await fetchSubscriptions(sid, tid);
            } else {
                // Multiple tenants OR single tenant needing re-auth, let user choose
                setStep('select-tenant');
            }
        } catch (err) {
            console.error('❌ Error fetching tenants:', err);
            await fetchSubscriptions(sid);
        }
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
                console.log('✅ Subscriptions:', subs);
                setSubscriptions(subs);
                setStep('select-subscriptions');
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

    const handleTenantSelect = async (tenantId) => {
        const tenant = tenants.find(t => t.tenant_id === tenantId);

        // Check if this tenant has subscriptions
        if (tenant && tenant.subscription_count === 0) {
            // This is a guest tenant that needs re-authentication
            console.log(`🔄 Tenant ${tenantId} has 0 subscriptions, triggering re-authentication...`);

            // Show confirmation to user
            if (window.confirm(
                `This tenant (${tenant.display_name}) requires re-authentication to access its subscriptions.\n\n` +
                `You will receive a new device code to authenticate specifically to this tenant.\n\n` +
                `Click OK to continue.`
            )) {
                // Trigger new authentication flow for this specific tenant
                await handleStartAuth(tenantId);
            }
        } else {
            // Tenant has subscriptions, proceed normally
            setSelectedTenant(tenantId);
            await fetchSubscriptions(sessionId, tenantId);
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

        navigate('/assessment', {
            state: {
                authMethod: 'az-cli',
                sessionId: sessionId,
                subscriptionIds: selectedSubscriptions,
                subscriptions: selectedSubDetails,
                tenantId: selectedTenant // Pass tenant ID so we can preserve state if user goes back
            }
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
                                    Click the button below to start the authentication process. You'll receive a code
                                    to enter on the Microsoft login page to securely authenticate with your Azure account.
                                </p>
                            </div>

                            <button
                                onClick={() => handleStartAuth()}
                                disabled={loading}
                                className="w-full bg-indigo-600 text-white py-3 px-6 rounded-lg font-semibold hover:bg-indigo-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                            >
                                {loading ? 'Starting...' : 'Start Authentication'}
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
                                        {deviceCodeInfo.userCode === 'Loading...' ? (
                                            <div className="flex items-center justify-center py-2">
                                                <div className="animate-pulse flex space-x-2">
                                                    <div className="h-4 w-4 bg-green-400 rounded-full"></div>
                                                    <div className="h-4 w-4 bg-green-400 rounded-full"></div>
                                                    <div className="h-4 w-4 bg-green-400 rounded-full"></div>
                                                </div>
                                            </div>
                                        ) : (
                                            <p className="text-3xl font-mono font-bold text-center text-gray-900 tracking-wider">
                                                {deviceCodeInfo.userCode}
                                            </p>
                                        )}
                                    </div>
                                    <button
                                        onClick={() => copyToClipboard(deviceCodeInfo.userCode)}
                                        disabled={deviceCodeInfo.userCode === 'Loading...'}
                                        className={`px-4 py-2 rounded-lg transition-colors ${deviceCodeInfo.userCode === 'Loading...'
                                            ? 'bg-gray-300 text-gray-500 cursor-not-allowed'
                                            : copied
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
                {step === 'loading-tenants' && (
                    <div className="bg-white rounded-2xl shadow-xl p-8">
                        <div className="text-center">
                            <div className="inline-block p-4 bg-green-100 rounded-full mb-4">
                                <svg className="w-12 h-12 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                                </svg>
                            </div>
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">Authentication Successful</h2>
                            <p className="text-gray-600 mb-6">Loading your tenants and subscriptions...</p>
                            <div className="inline-flex items-center gap-2 text-gray-600">
                                <div className="animate-spin h-5 w-5 border-2 border-indigo-600 border-t-transparent rounded-full"></div>
                                <span>Please wait</span>
                            </div>
                        </div>
                    </div>
                )}

                {/* Tenant Selection Step */}
                {step === 'select-tenant' && (
                    <div className="bg-white rounded-2xl shadow-xl p-8">
                        <div className="text-center mb-8">
                            <div className="inline-block p-4 bg-blue-100 rounded-full mb-4">
                                <svg className="w-12 h-12 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4" />
                                </svg>
                            </div>
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">Select Tenant</h2>
                            <p className="text-gray-600">Please choose which organization's resources you want to assess</p>
                        </div>

                        {error && (
                            <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700">
                                {error}
                            </div>
                        )}

                        {loading && (
                            <div className="text-center mb-4">
                                <div className="inline-flex items-center gap-2 text-indigo-600">
                                    <div className="animate-spin h-5 w-5 border-2 border-indigo-600 border-t-transparent rounded-full"></div>
                                    <span>Fetching subscriptions...</span>
                                </div>
                            </div>
                        )}

                        <div className="space-y-4">
                            {tenants.map((tenant) => (
                                <button
                                    key={tenant.tenant_id}
                                    onClick={() => handleTenantSelect(tenant.tenant_id)}
                                    className="w-full flex items-center justify-between p-4 bg-white border-2 border-gray-200 rounded-xl hover:border-blue-500 hover:bg-blue-50 transition-all group"
                                >
                                    <div className="flex flex-col items-start">
                                        <span className="font-semibold text-gray-900 group-hover:text-blue-700">{tenant.display_name}</span>
                                        <span className="text-sm text-gray-500 group-hover:text-blue-600">{tenant.tenant_id}</span>
                                    </div>
                                    <div className="flex items-center gap-3">
                                        <div className="text-right mr-2">
                                            {tenant.subscription_count > 0 ? (
                                                <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800 border border-green-200">
                                                    {tenant.subscription_count} Subscription(s)
                                                </span>
                                            ) : (
                                                <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-800 border border-amber-200">
                                                    Action Required
                                                </span>
                                            )}
                                        </div>
                                        <svg className="w-5 h-5 text-gray-400 group-hover:text-blue-500 transition-colors" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                                        </svg>
                                    </div>
                                </button>
                            ))}
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

                            {/* Option to go back to tenants if multiple tenants exist */}
                            {tenants.length > 1 && (
                                <button
                                    onClick={() => setStep('select-tenant')}
                                    className="mt-2 text-sm text-indigo-600 hover:text-indigo-800 underline"
                                >
                                    ← Change Tenant ({tenants.find(t => t.tenant_id === selectedTenant)?.display_name || 'Selected'})
                                </button>
                            )}
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

                {/* Completed Step */}
                {step === 'completed' && (
                    <div className="bg-white rounded-2xl shadow-xl p-8 text-center">
                        <div className="inline-block p-4 bg-green-100 rounded-full mb-4">
                            <svg className="w-12 h-12 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                            </svg>
                        </div>
                        <h2 className="text-2xl font-bold text-gray-900 mb-2">Authentication Successful!</h2>
                        <p className="text-gray-600">Redirecting to assessment page...</p>
                    </div>
                )}

            </div>
        </div >
    );
}


