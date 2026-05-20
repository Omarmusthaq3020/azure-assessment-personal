import { useEffect, useState } from 'react';
import axios from 'axios';

export default function GcpOAuthCallbackPage() {
    const [status, setStatus] = useState('Completing Google sign-in...');

    useEffect(() => {
        const completeOAuth = async () => {
            const params = new URLSearchParams(window.location.search);
            const code = params.get('code');
            const state = params.get('state');
            const error = params.get('error');

            if (error) {
                setStatus(`Google sign-in failed: ${error}`);
                if (window.opener) {
                    window.opener.postMessage({ type: 'gcp-oauth-error', error }, window.location.origin);
                }
                return;
            }

            if (!code || !state) {
                setStatus('Missing OAuth code/state.');
                return;
            }

            try {
                await axios.post('/api/auth/gcp/oauth-exchange', {
                    session_id: state,
                    state,
                    code
                });

                setStatus('Google sign-in complete. You can close this window.');
                if (window.opener) {
                    window.opener.postMessage({ type: 'gcp-oauth-success', sessionId: state }, window.location.origin);
                }
                setTimeout(() => window.close(), 400);
            } catch (err) {
                const detail = err.response?.data?.detail || 'OAuth exchange failed';
                setStatus(detail);
                if (window.opener) {
                    window.opener.postMessage({ type: 'gcp-oauth-error', error: detail }, window.location.origin);
                }
            }
        };

        completeOAuth();
    }, []);

    return (
        <div className="min-h-screen bg-gray-900 text-gray-100 flex items-center justify-center p-6">
            <div className="max-w-lg w-full bg-gray-800 border border-gray-700 rounded-xl p-6 text-center">
                <h1 className="text-xl font-semibold mb-3">GCP Authentication</h1>
                <p className="text-gray-300">{status}</p>
            </div>
        </div>
    );
}
