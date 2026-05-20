import { Routes, Route, useNavigate } from 'react-router-dom';
import PreAssessmentQuestionnaire from './components/PreAssessmentQuestionnaire';
import AzCliDeviceCodeLogin from './pages/AzCliDeviceCodeLogin';
import AwsLoginPage from './pages/AwsLoginPage';
import GcpLoginPage from './pages/GcpLoginPage';
import GcpOAuthCallbackPage from './pages/GcpOAuthCallbackPage';
import AssessmentPage from './pages/AssessmentPage';
import DiagramPlaygroundPage from './pages/DiagramPlaygroundPage';
import { useState } from 'react';

function App() {
  const [assessmentConfig, setAssessmentConfig] = useState(null);
  const navigate = useNavigate();

  const handleConfigSubmit = (config) => {
    setAssessmentConfig(config);

    // Route based on migration type
    if (config.migration_type === 'azure_to_azure') {
      navigate('/azure-login');
    } else if (config.migration_type === 'aws_to_azure') {
      navigate('/aws-login');
    } else if (config.migration_type === 'gcp_to_azure') {
      navigate('/gcp-login');
    }
  };

  return (
    <Routes>
      <Route path="/" element={<PreAssessmentQuestionnaire onSubmit={handleConfigSubmit} />} />
      <Route path="/azure-login" element={<AzCliDeviceCodeLogin assessmentConfig={assessmentConfig} />} />
      <Route path="/aws-login" element={<AwsLoginPage assessmentConfig={assessmentConfig} />} />
      <Route path="/gcp-login" element={<GcpLoginPage assessmentConfig={assessmentConfig} />} />
      <Route path="/gcp-oauth-callback" element={<GcpOAuthCallbackPage />} />
      <Route path="/assessment" element={<AssessmentPage />} />
      <Route path="/diagram-playground" element={<DiagramPlaygroundPage />} />
    </Routes>
  );
}

export default App;
