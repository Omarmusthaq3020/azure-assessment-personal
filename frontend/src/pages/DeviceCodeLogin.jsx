import { useState } from 'react';
import { useNavigate } from 'react-router-dom';

export default function DeviceCodeLogin() {
    const [step, setStep] = useState('tenant'); // 'tenant', 'device-code', 'waiting', 'subscriptions'
    const [tenantId, setTenantId] = useState('');
    const [sessionId, setSessionId] = useState('');
    const [deviceCodeInfo, setDeviceCodeInfo] = useState(null);
    const [subscriptions, setSubscriptions] = useState([]);
    const [selectedSubs, setSelectedSubs] = useState([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState('');

    const navigate = useNavigate();

    const handleStartAuth = async () => {
        if (!tenantId.trim()) {
            setError('Please enter a Tenant ID');
            return;
        }

        setLoading(true);
        setError('');

        try {
            const response = await fetch('/api/auth/start', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ tenant_id: tenantId })
            });

            if (!response.ok) {
                throw new Error('Failed to start authentication');
            }

            const data = await response.json();
            setSessionId(data.session_id);
            setDeviceCodeInfo({
                userCode: data.user_code,
                verificationUrl: data.verification_url,
                message: data.message,
                expiresAt: data.expires_at
            });
            setStep('device-code');

            // Start polling for authentication status
            pollAuthStatus(data.session_id);

        } catch (err) {
            setError(err.message);
        } finally {
            setLoading(false);
        }
    };

    const pollAuthStatus = async (sid) => {
        const pollInterval = setInterval(async () => {
            try {
                const response = await fetch(`/api/auth/status/${sid}`);
                const data = await response.json();

                if (data.status === 'authenticated') {
                    clearInterval(pollInterval);
                    setStep('waiting');
                    await loadSubscriptions(sid);
                } else if (data.status === 'failed') {
                    clearInterval(pollInterval);
                    setError(data.error || 'Authentication failed');
                    setStep('tenant');
                }
            } catch (err) {
                console.error('Polling error:', err);
            }
        }, 3000); // Poll every 3 seconds

        // Stop polling after 15 minutes
        setTimeout(() => clearInterval(pollInterval), 900000);
    };

    const loadSubscriptions = async (sid) => {
        try {
            const response = await fetch(`/api/auth/subscriptions/${sid}`);
            if (!response.ok) throw new Error('Failed to load subscriptions');

            const data = await response.json();
            setSubscriptions(data);
            setSelectedSubs(data.map(s => s.subscription_id)); // Select all by default
            setStep('subscriptions');
        } catch (err) {
            setError(err.message);
        }
    };

    const handleRunAssessment = () => {
        // Navigate to assessment page with session and selected subscriptions
        navigate('/assessment', {
            state: {
                sessionId,
                subscriptionIds: selectedSubs
            }
        });
    };

    const copyToClipboard = (text) => {
        navigator.clipboard.writeText(text);
    };

    return (
        <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-center justify-center p-4">
            <div className="max-w-2xl w-full">

                {/* Tenant ID Input Step */}
                {step === 'tenant' && (
                    <div className="bg-white rounded-2xl shadow-xl p-8">
                        <div className="text-center mb-8">
                            <div className="inline-block p-4 bg-blue-100 rounded-full mb-4">
                                <svg className="w-12 h-12 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
                                </svg>
                            </div>
                            <h1 className="text-3xl font-bold text-gray-900 mb-2">Azure Migration Assessment</h1>
                            <p className="text-gray-600">Enter your Azure Tenant ID to begin</p>
                        </div>

                        {error && (
                            <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700">
                                {error}
                            </div>
                        )}

                        <div className="space-y-6">
                            <div>
                                <label className="block text-sm font-medium text-gray-700 mb-2">
                                    Azure Tenant ID
                                </label>
                                <input
                                    type="text"
                                    value={tenantId}
                                    onChange={(e) => setTenantId(e.target.value)}
                                    placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                                />
                                <p className="mt-2 text-sm text-gray-500">
                                    Find this in: Azure Portal → Azure Active Directory → Overview
                                </p>
                            </div>

                            <button
                                onClick={handleStartAuth}
                                disabled={loading}
                                className="w-full bg-blue-600 text-white py-3 px-6 rounded-lg font-semibold hover:bg-blue-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                            >
                                {loading ? 'Starting...' : 'Start Authentication'}
                            </button>

                            <div className="text-center">
                                <button
                                    onClick={() => navigate('/login-cli')}
                                    className="text-sm text-gray-600 hover:text-gray-900 underline"
                                >
                                    Use Azure CLI authentication instead
                                </button>
                            </div>
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
                                        className="px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
                                    >
                                        Copy
                                    </button>
                                </div>
                            </div>

                            <div className="text-center">
                                <div className="inline-flex items-center gap-2 text-gray-600">
                                    <div className="animate-spin h-5 w-5 border-2 border-blue-600 border-t-transparent rounded-full"></div>
                                    <span>Waiting for you to complete authentication...</span>
                                </div>
                            </div>
                        </div>
                    </div>
                )}

                {/* Loading Step */}
                {step === 'waiting' && (
                    <div className="bg-white rounded-2xl shadow-xl p-8 text-center">
                        <div className="animate-spin h-16 w-16 border-4 border-blue-600 border-t-transparent rounded-full mx-auto mb-4"></div>
                        <h2 className="text-xl font-semibold text-gray-900 mb-2">Loading Subscriptions...</h2>
                        <p className="text-gray-600">Please wait while we fetch your Azure subscriptions</p>
                    </div>
                )}

                {/* Subscription Selection Step */}
                {step === 'subscriptions' && (
                    <div className="bg-white rounded-2xl shadow-xl p-8">
                        <div className="text-center mb-8">
                            <div className="inline-block p-4 bg-purple-100 rounded-full mb-4">
                                <svg className="w-12 h-12 text-purple-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" />
                                </svg>
                            </div>
                            <h2 className="text-2xl font-bold text-gray-900 mb-2">Select Subscriptions</h2>
                            <p className="text-gray-600">Choose which subscriptions to assess</p>
                        </div>

                        <div className="space-y-4 mb-6">
                            {subscriptions.map((sub) => (
                                <label
                                    key={sub.subscription_id}
                                    className="flex items-center p-4 border border-gray-200 rounded-lg hover:bg-gray-50 cursor-pointer transition-colors"
                                >
                                    <input
                                        type="checkbox"
                                        checked={selectedSubs.includes(sub.subscription_id)}
                                        onChange={(e) => {
                                            if (e.target.checked) {
                                                setSelectedSubs([...selectedSubs, sub.subscription_id]);
                                            } else {
                                                setSelectedSubs(selectedSubs.filter(id => id !== sub.subscription_id));
                                            }
                                        }}
                                        className="w-5 h-5 text-blue-600 rounded focus:ring-2 focus:ring-blue-500"
                                    />
                                    <div className="ml-4 flex-1">
                                        <p className="font-medium text-gray-900">{sub.display_name}</p>
                                        <p className="text-sm text-gray-500">{sub.subscription_id}</p>
                                    </div>
                                    <span className={`px-3 py-1 rounded-full text-xs font-medium ${sub.state === 'Enabled' ? 'bg-green-100 text-green-800' : 'bg-gray-100 text-gray-800'
                                        }`}>
                                        {sub.state}
                                    </span>
                                </label>
                            ))}
                        </div>

                        <button
                            onClick={handleRunAssessment}
                            disabled={selectedSubs.length === 0}
                            className="w-full bg-purple-600 text-white py-3 px-6 rounded-lg font-semibold hover:bg-purple-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                        >
                            Run Assessment ({selectedSubs.length} selected)
                        </button>
                    </div>
                )}

            </div>
        </div>
    );
}
