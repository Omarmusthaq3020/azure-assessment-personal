import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import axios from 'axios';
import DrawioViewer from '../components/DrawioViewer';

export default function DiagramPlaygroundPage() {
    const navigate = useNavigate();
    const [selectedFile, setSelectedFile] = useState(null);
    const [diagramType, setDiagramType] = useState('network');
    const [cloudType, setCloudType] = useState('azure');
    const [dragOver, setDragOver] = useState(false);
    const [loading, setLoading] = useState(false);
    const [drawioXml, setDrawioXml] = useState('');
    const [error, setError] = useState('');

    const onFilePicked = (file) => {
        if (!file) return;
        if (!file.name.toLowerCase().endsWith('.xlsx')) {
            setError('Please upload an .xlsx assessment report file.');
            return;
        }
        setSelectedFile(file);
        setError('');
        setDrawioXml('');
    };

    const handleDrop = (e) => {
        e.preventDefault();
        setDragOver(false);
        const file = e.dataTransfer.files?.[0];
        onFilePicked(file);
    };

    const handleGenerate = async () => {
        if (!selectedFile) {
            setError('Please choose an Excel file first.');
            return;
        }

        setLoading(true);
        setError('');
        setDrawioXml('');
        try {
            const form = new FormData();
            form.append('file', selectedFile);
            form.append('diagram_type', diagramType);
            form.append('cloud_type', cloudType);

            const response = await axios.post('/api/generate-architecture-upload', form, {
                headers: { 'Content-Type': 'multipart/form-data' }
            });
            setDrawioXml(response.data.drawio_xml || '');
        } catch (err) {
            setError(err.response?.data?.detail || 'Failed to generate diagram.');
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="min-h-screen bg-gray-900 text-gray-100 p-8">
            <div className="max-w-5xl mx-auto">
                <div className="flex items-center justify-between mb-8">
                    <div>
                        <h1 className="text-3xl font-bold text-white">Diagram Playground</h1>
                        <p className="text-gray-400 mt-1">Upload an assessment Excel and generate diagram without running assessment.</p>
                    </div>
                    <button
                        onClick={() => navigate('/assessment')}
                        className="px-4 py-2 bg-gray-700 hover:bg-gray-600 text-white rounded-lg border border-gray-600"
                    >
                        Back to Assessment
                    </button>
                </div>

                <div className="bg-gray-800 border border-gray-700 rounded-xl p-6 mb-6">
                    <div
                        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                        onDragLeave={() => setDragOver(false)}
                        onDrop={handleDrop}
                        className={`border-2 border-dashed rounded-lg p-8 text-center transition ${
                            dragOver ? 'border-blue-400 bg-blue-900/20' : 'border-gray-600 bg-gray-900/40'
                        }`}
                    >
                        <input
                            type="file"
                            accept=".xlsx"
                            className="hidden"
                            id="diagram-file-input"
                            onChange={(e) => onFilePicked(e.target.files?.[0])}
                        />
                        <label htmlFor="diagram-file-input" className="cursor-pointer text-blue-300 hover:text-blue-200">
                            Click to choose .xlsx file
                        </label>
                        <p className="text-gray-400 text-sm mt-2">or drag and drop here</p>
                        {selectedFile && (
                            <p className="text-green-400 text-sm mt-3">Selected: {selectedFile.name}</p>
                        )}
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mt-5">
                        <div>
                            <label className="block text-sm text-gray-300 mb-1">Diagram Type</label>
                            <select
                                value={diagramType}
                                onChange={(e) => setDiagramType(e.target.value)}
                                className="w-full px-3 py-2 bg-gray-700 border border-gray-600 rounded-lg text-white"
                            >
                                <option value="network">Complete Architecture (with peering)</option>
                                <option value="resource_groups">Resource Group wise</option>
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm text-gray-300 mb-1">Cloud Type</label>
                            <select
                                value={cloudType}
                                onChange={(e) => setCloudType(e.target.value)}
                                className="w-full px-3 py-2 bg-gray-700 border border-gray-600 rounded-lg text-white"
                            >
                                <option value="azure">Azure</option>
                                <option value="aws">AWS</option>
                                <option value="gcp">GCP</option>
                            </select>
                        </div>
                        <div className="flex items-end">
                            <button
                                onClick={handleGenerate}
                                disabled={loading || !selectedFile}
                                className={`w-full py-2 rounded-lg text-white font-medium ${
                                    loading || !selectedFile ? 'bg-emerald-900 cursor-not-allowed' : 'bg-emerald-600 hover:bg-emerald-500'
                                }`}
                            >
                                {loading ? 'Generating...' : 'Generate Diagram'}
                            </button>
                        </div>
                    </div>

                    {error && (
                        <div className="mt-4 p-3 text-sm text-red-200 bg-red-900/40 border border-red-800 rounded-lg">
                            {error}
                        </div>
                    )}
                </div>

                {drawioXml && (
                    <div className="bg-gray-800 border border-gray-700 rounded-xl p-6">
                        <h2 className="text-xl font-semibold text-white mb-4">Generated Diagram</h2>
                        <DrawioViewer xml={drawioXml} />
                    </div>
                )}
            </div>
        </div>
    );
}
