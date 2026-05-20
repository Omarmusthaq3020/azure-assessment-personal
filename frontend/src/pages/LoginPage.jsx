import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import axios from 'axios';
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../authConfig";

export default function LoginPage() {
    const [authType, setAuthType] = useState('service_principal'); // 'service_principal' or 'azure_ad'
    const [formData, setFormData] = useState({
        tenant_id: '',
        client_id: '',
        client_secret: ''
    });
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState('');
    const navigate = useNavigate();
    const { instance } = useMsal();

    const handleChange = (e) => {
        setFormData({ ...formData, [e.target.name]: e.target.value });
    };

    const handleServicePrincipalLogin = async (e) => {
        e.preventDefault();
        setLoading(true);
        setError('');

        try {
            await axios.post('/api/login', {
                auth_type: 'service_principal',
                ...formData
            });
            // Clear any stored token
            localStorage.removeItem('azure_access_token');
            navigate('/assessment');
        } catch (err) {
            setError(err.response?.data?.detail || 'Login failed. Please check your credentials.');
        } finally {
            setLoading(false);
        }
    };

    const handleAzureADLogin = async () => {
        setLoading(true);
        setError('');
        try {
            const loginResponse = await instance.loginPopup(loginRequest);
            // Store token for use in assessment calls
            localStorage.setItem('azure_access_token', loginResponse.accessToken);
            navigate('/assessment');
        } catch (err) {
            console.error(err);
            setError('Azure AD Login failed: ' + err.message);
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="flex items-center justify-center min-h-screen bg-gray-900">
            <div className="w-full max-w-md p-8 space-y-6 bg-gray-800 rounded-xl shadow-2xl border border-gray-700">
                <h2 className="text-3xl font-bold text-center text-white">Azure Migration Assessment</h2>

                {/* Auth Type Tabs */}
                <div className="flex border-b border-gray-700">
                    <button
                        className={`flex-1 py-2 text-sm font-medium ${authType === 'service_principal' ? 'text-blue-500 border-b-2 border-blue-500' : 'text-gray-400 hover:text-gray-300'}`}
                        onClick={() => setAuthType('service_principal')}
                    >
                        Service Principal
                    </button>
                    <button
                        className={`flex-1 py-2 text-sm font-medium ${authType === 'azure_ad' ? 'text-blue-500 border-b-2 border-blue-500' : 'text-gray-400 hover:text-gray-300'}`}
                        onClick={() => setAuthType('azure_ad')}
                    >
                        Azure AD (Interactive)
                    </button>
                </div>

                {error && (
                    <div className="p-4 text-sm text-red-200 bg-red-900/50 border border-red-800 rounded-lg">
                        {error}
                    </div>
                )}

                {authType === 'service_principal' ? (
                    <form onSubmit={handleServicePrincipalLogin} className="space-y-5">
                        <div>
                            <label className="block mb-2 text-sm font-medium text-gray-300">Tenant ID</label>
                            <input
                                type="text"
                                name="tenant_id"
                                required
                                className="w-full px-4 py-3 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-white placeholder-gray-500 transition"
                                placeholder="Enter Tenant ID"
                                value={formData.tenant_id}
                                onChange={handleChange}
                            />
                        </div>
                        <div>
                            <label className="block mb-2 text-sm font-medium text-gray-300">Client ID</label>
                            <input
                                type="text"
                                name="client_id"
                                required
                                className="w-full px-4 py-3 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-white placeholder-gray-500 transition"
                                placeholder="Enter Client ID"
                                value={formData.client_id}
                                onChange={handleChange}
                            />
                        </div>
                        <div>
                            <label className="block mb-2 text-sm font-medium text-gray-300">Client Secret</label>
                            <input
                                type="password"
                                name="client_secret"
                                required
                                className="w-full px-4 py-3 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-white placeholder-gray-500 transition"
                                placeholder="Enter Client Secret"
                                value={formData.client_secret}
                                onChange={handleChange}
                            />
                        </div>

                        <button
                            type="submit"
                            disabled={loading}
                            className={`w-full py-3 text-white font-semibold rounded-lg shadow-lg transition duration-200 
                  ${loading ? 'bg-blue-700 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-500 hover:shadow-blue-500/30'}`}
                        >
                            {loading ? 'Authenticating...' : 'Login'}
                        </button>
                    </form>
                ) : (
                    <div className="space-y-5 text-center">
                        <p className="text-gray-400 text-sm">
                            Login with your Azure Portal credentials. A popup window will appear.
                        </p>
                        <button
                            onClick={handleAzureADLogin}
                            disabled={loading}
                            className={`w-full py-3 text-white font-semibold rounded-lg shadow-lg transition duration-200 
                  ${loading ? 'bg-blue-700 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-500 hover:shadow-blue-500/30'}`}
                        >
                            {loading ? 'Waiting for Popup...' : 'Login with Azure AD'}
                        </button>
                    </div>
                )}
            </div>
        </div>
    );
}
