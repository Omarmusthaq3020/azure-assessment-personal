import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import axios from 'axios';

export default function AwsLoginPage() {
    const AWS_REGIONS = [
        { id: 'us-east-1', name: 'us-east-1 (N. Virginia)' },
        { id: 'us-east-2', name: 'us-east-2 (Ohio)' },
        { id: 'us-west-1', name: 'us-west-1 (N. California)' },
        { id: 'us-west-2', name: 'us-west-2 (Oregon)' },
        { id: 'af-south-1', name: 'af-south-1 (Cape Town)' },
        { id: 'ap-east-1', name: 'ap-east-1 (Hong Kong)' },
        { id: 'ap-south-1', name: 'ap-south-1 (Mumbai)' },
        { id: 'ap-south-2', name: 'ap-south-2 (Hyderabad)' },
        { id: 'ap-northeast-3', name: 'ap-northeast-3 (Osaka)' },
        { id: 'ap-northeast-2', name: 'ap-northeast-2 (Seoul)' },
        { id: 'ap-southeast-1', name: 'ap-southeast-1 (Singapore)' },
        { id: 'ap-southeast-2', name: 'ap-southeast-2 (Sydney)' },
        { id: 'ap-southeast-3', name: 'ap-southeast-3 (Jakarta)' },
        { id: 'ap-northeast-1', name: 'ap-northeast-1 (Tokyo)' },
        { id: 'ca-central-1', name: 'ca-central-1 (Central)' },
        { id: 'eu-central-1', name: 'eu-central-1 (Frankfurt)' },
        { id: 'eu-west-1', name: 'eu-west-1 (Ireland)' },
        { id: 'eu-west-2', name: 'eu-west-2 (London)' },
        { id: 'eu-south-1', name: 'eu-south-1 (Milan)' },
        { id: 'eu-west-3', name: 'eu-west-3 (Paris)' },
        { id: 'eu-south-2', name: 'eu-south-2 (Spain)' },
        { id: 'eu-north-1', name: 'eu-north-1 (Stockholm)' },
        { id: 'eu-central-2', name: 'eu-central-2 (Zurich)' },
        { id: 'me-south-1', name: 'me-south-1 (Bahrain)' },
        { id: 'me-central-1', name: 'me-central-1 (UAE)' },
        { id: 'sa-east-1', name: 'sa-east-1 (São Paulo)' }
    ];

    const [formData, setFormData] = useState({
        aws_access_key_id: '',
        aws_secret_access_key: '',
        aws_session_token: '',
        region_names: ['us-east-1'] // Default to us-east-1
    });
    const [loading, setLoading] = useState(false);
    const [discoveryLoading, setDiscoveryLoading] = useState(false);
    const [error, setError] = useState('');
    const navigate = useNavigate();

    const handleChange = (e) => {
        setFormData({ ...formData, [e.target.name]: e.target.value });
    };

    const handleDiscoverRegions = async () => {
        if (!formData.aws_access_key_id || !formData.aws_secret_access_key) {
            setError('Please enter your Access Key ID and Secret Access Key first.');
            return;
        }

        setDiscoveryLoading(true);
        setError('');

        try {
            const response = await axios.post('/api/auth/aws/discover-regions', {
                ...formData,
                region_name: formData.region_names[0] || 'us-east-1' // Use first selected or fallback
            });
            
            const discovered = response.data; // List of {id, has_resources, resource_count}
            if (discovered.length === 0) {
                setError('No active resources found in any region. Please select regions manually or check permissions.');
            } else {
                const discoveredIds = discovered.map(r => r.id);
                setFormData(prev => ({
                    ...prev,
                    region_names: discoveredIds
                }));
                // Optionally show a message
                console.log(`Discovered ${discoveredIds.length} active regions.`);
            }
        } catch (err) {
            setError(err.response?.data?.detail || 'Region discovery failed. Please check your credentials and permissions.');
        } finally {
            setDiscoveryLoading(false);
        }
    };

    const handleRegionToggle = (regionId) => {
        setFormData(prev => {
            const current = prev.region_names;
            if (current.includes(regionId)) {
                return { ...prev, region_names: current.filter(id => id !== regionId) };
            } else {
                return { ...prev, region_names: [...current, regionId] };
            }
        });
    };

    const handleSelectAllRegions = (checked) => {
        if (checked) {
            setFormData(prev => ({ ...prev, region_names: AWS_REGIONS.map(r => r.id) }));
        } else {
            setFormData(prev => ({ ...prev, region_names: [] }));
        }
    };

    const handleAwsLogin = async (e) => {
        e.preventDefault();
        if (formData.region_names.length === 0) {
            setError('Please select at least one region to scan.');
            return;
        }

        setLoading(true);
        setError('');

        try {
            // Validate credentials with backend (use the first region for validation)
            await axios.post('/api/auth/aws/validate', {
                ...formData,
                region_name: formData.region_names[0]
            });

            navigate('/assessment', { state: { awsCredentials: formData } });
        } catch (err) {
            setError(err.response?.data?.detail || 'AWS Authentication failed. Please check your credentials.');
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="flex items-center justify-center min-h-screen bg-gray-900 overflow-y-auto py-12">
            <div className="w-full max-w-2xl p-8 space-y-6 bg-gray-800 rounded-xl shadow-2xl border border-gray-700 mx-4">
                <div className="text-center">
                    <h2 className="text-3xl font-bold text-white">AWS Migration Assessment</h2>
                    <p className="mt-2 text-gray-400">Scan multiple regions in parallel for a unified report.</p>
                </div>

                {error && (
                    <div className="p-4 text-sm text-red-200 bg-red-900/50 border border-red-800 rounded-lg animate-in fade-in slide-in-from-top-2">
                        {error}
                    </div>
                )}

                <form onSubmit={handleAwsLogin} className="space-y-6">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div className="md:col-span-2">
                             <label className="block mb-1 text-sm font-medium text-gray-300">Access Key ID</label>
                            <input
                                type="text"
                                name="aws_access_key_id"
                                required
                                className="w-full px-4 py-2 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 text-white placeholder-gray-500"
                                placeholder="AKIA..."
                                value={formData.aws_access_key_id}
                                onChange={handleChange}
                            />
                        </div>
                        <div className="md:col-span-2">
                            <label className="block mb-1 text-sm font-medium text-gray-300">Secret Access Key</label>
                            <input
                                type="password"
                                name="aws_secret_access_key"
                                required
                                className="w-full px-4 py-2 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 text-white placeholder-gray-500"
                                placeholder="Secret Key"
                                value={formData.aws_secret_access_key}
                                onChange={handleChange}
                            />
                        </div>
                        <div className="md:col-span-2">
                            <label className="block mb-1 text-sm font-medium text-gray-300">
                                Session Token <span className="text-xs text-gray-500">(Optional)</span>
                            </label>
                            <textarea
                                name="aws_session_token"
                                rows="2"
                                className="w-full px-4 py-2 bg-gray-700 border border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 text-white placeholder-gray-500 text-xs font-mono"
                                placeholder="Required for temporary credentials..."
                                value={formData.aws_session_token}
                                onChange={handleChange}
                            />
                        </div>
                    </div>

                    <div className="pt-2">
                        <button
                            type="button"
                            onClick={handleDiscoverRegions}
                            disabled={discoveryLoading || !formData.aws_access_key_id}
                            className={`w-full py-2 bg-blue-600/20 hover:bg-blue-600/30 border border-blue-500/50 text-blue-400 text-sm font-semibold rounded-lg transition-all flex items-center justify-center
                                ${discoveryLoading ? 'opacity-50 cursor-not-allowed' : ''}
                                ${!formData.aws_access_key_id ? 'opacity-20 cursor-not-allowed grayscale' : ''}`}
                        >
                            {discoveryLoading ? (
                                <>
                                    <svg className="w-4 h-4 mr-2 animate-spin" viewBox="0 0 24 24">
                                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                                    </svg>
                                    Discovering Active Regions...
                                </>
                            ) : (
                                <>
                                    <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                                    </svg>
                                    Discover Active Regions
                                </>
                            )}
                        </button>
                        <p className="mt-2 text-[10px] text-gray-400 text-center italic">
                             Automatically finds regions where you currently have resources.
                        </p>
                    </div>

                    <div>
                        <div className="flex justify-between items-center mb-4">
                            <label className="text-sm font-medium text-gray-300">Regions to Scan</label>
                            <label className="flex items-center text-xs text-blue-400 hover:text-blue-300 cursor-pointer">
                                <input 
                                    type="checkbox" 
                                    className="mr-2"
                                    onChange={(e) => handleSelectAllRegions(e.target.checked)}
                                    checked={formData.region_names.length === AWS_REGIONS.length}
                                />
                                Select All
                            </label>
                        </div>
                        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2 max-h-48 overflow-y-auto bg-gray-900/50 p-4 rounded-lg border border-gray-700">
                            {AWS_REGIONS.map(region => (
                                <label key={region.id} className={`flex items-center space-x-2 text-xs cursor-pointer p-1 rounded-md transition-colors
                                    ${formData.region_names.includes(region.id) ? 'bg-blue-500/10 text-white' : 'text-gray-400 hover:text-white'}`}>
                                    <input
                                        type="checkbox"
                                        checked={formData.region_names.includes(region.id)}
                                        onChange={() => handleRegionToggle(region.id)}
                                        className="rounded border-gray-600 text-blue-600 focus:ring-blue-500"
                                    />
                                    <span>{region.name}</span>
                                </label>
                            ))}
                        </div>
                    </div>

                    <button
                        type="submit"
                        disabled={loading || formData.region_names.length === 0}
                        className={`w-full py-4 text-white font-bold rounded-xl shadow-lg transition-all
                  ${loading ? 'bg-orange-800 cursor-not-allowed' : 'bg-orange-600 hover:bg-orange-500 hover:shadow-orange-500/40'}
                  ${formData.region_names.length === 0 ? 'opacity-50 grayscale' : ''}`}
                    >
                        {loading ? (
                            <div className="flex items-center justify-center">
                                <svg className="w-5 h-5 mr-3 animate-spin" viewBox="0 0 24 24">
                                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                                </svg>
                                Running Parallel Checks...
                            </div>
                        ) : 'Start Parallel Assessment'}
                    </button>
                    <p className="text-center text-xs text-gray-500">
                        The assessment will extract resources from all selected regions in one consolidated report.
                    </p>
                </form>
            </div>
        </div>
    );
}
