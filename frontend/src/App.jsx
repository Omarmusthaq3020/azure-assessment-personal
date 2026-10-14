import { Routes, Route, useNavigate } from 'react-router-dom';
import PreAssessmentQuestionnaire from './components/PreAssessmentQuestionnaire';
import AzCliDeviceCodeLogin from './pages/AzCliDeviceCodeLogin';
import AwsLoginPage from './pages/AwsLoginPage';
import GcpLoginPage from './pages/GcpLoginPage';
import GcpOAuthCallbackPage from './pages/GcpOAuthCallbackPage';
import AssessmentPage from './pages/AssessmentPage';
import DiagramPlaygroundPage from './pages/DiagramPlaygroundPage';

function App() {
  const navigate = useNavigate();

  const handleConfigSubmit = (config) => {
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
      <Route path="/azure-login" element={<AzCliDeviceCodeLogin />} />
      <Route path="/aws-login" element={<AwsLoginPage />} />
      <Route path="/gcp-login" element={<GcpLoginPage />} />
      <Route path="/gcp-oauth-callback" element={<GcpOAuthCallbackPage />} />
      <Route path="/assessment" element={<AssessmentPage />} />
      <Route path="/diagram-playground" element={<DiagramPlaygroundPage />} />
    </Routes>
  );
}

export default App;
