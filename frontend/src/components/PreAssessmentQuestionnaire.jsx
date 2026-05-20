import React, { useState } from 'react';
import './PreAssessmentQuestionnaire.css';

const PreAssessmentQuestionnaire = ({ onComplete, onSubmit }) => {
    const [currentQuestion, setCurrentQuestion] = useState(0);
    const [answers, setAnswers] = useState({});
    const [showResult, setShowResult] = useState(false);
    const [isEligible, setIsEligible] = useState(false);
    const [failureReasons, setFailureReasons] = useState([]);

    const questions = [
        {
            id: 'migration_type',
            question: 'What type of migration are you planning?',
            type: 'single',
            options: [
                { value: 'azure_to_azure', label: 'Azure to Azure (Microsoft Sponsorship to EA)', eligible: true },
                { value: 'aws_to_azure', label: 'AWS to Azure', eligible: true },
                { value: 'gcp_to_azure', label: 'GCP to Azure', eligible: true },
                { value: 'other_cloud_to_azure', label: 'Other Cloud to Azure', eligible: false, reason: 'This tool currently supports Azure-to-Azure, AWS-to-Azure, and GCP-to-Azure migrations only.' }
            ]
        },
        {
            id: 'azure_access',
            question: 'Do you have access to the source Azure subscription?',
            type: 'single',
            options: [
                { value: 'yes_owner', label: 'Yes, I have Owner or Contributor access', eligible: true },
                { value: 'yes_reader', label: 'Yes, I have Reader access', eligible: true },
                { value: 'no', label: 'No, I don\'t have access', eligible: false, reason: 'You need at least Reader access to the source subscription to run this assessment.' }
            ]
        },
        {
            id: 'assessment_goal',
            question: 'What is your primary goal for this assessment?',
            type: 'single',
            options: [
                { value: 'move_feasibility', label: 'Determine which resources can be moved', eligible: true },
                { value: 'dependency_analysis', label: 'Understand resource dependencies', eligible: true },
                { value: 'cost_estimation', label: 'Get cost estimates for migration', eligible: false, reason: 'This tool does not provide cost estimation. It focuses on move feasibility and dependencies.' },
                { value: 'performance_analysis', label: 'Analyze performance impact', eligible: false, reason: 'This tool does not analyze performance. It focuses on move feasibility and dependencies.' }
            ]
        },
        {
            id: 'reservation_limitation',
            question: 'This tool cannot capture Azure Reservations data. Are you okay to proceed without reservation information?',
            type: 'single',
            options: [
                { value: 'yes', label: 'Yes, I understand and want to proceed', eligible: true },
                { value: 'no', label: 'No, I need reservation data', eligible: false, reason: 'This tool does not capture Azure Reservations. Please use Azure Portal or Azure CLI to view your reservations separately.' }
            ]
        },
        {
            id: 'resource_count',
            question: 'Approximately how many resources are in your subscription(s)?',
            type: 'single',
            options: [
                { value: 'small', label: 'Less than 100 resources', eligible: true },
                { value: 'medium', label: '100-500 resources', eligible: true },
                { value: 'large', label: '500-2000 resources', eligible: true },
                { value: 'very_large', label: 'More than 2000 resources', eligible: true, warning: 'Large environments may take 30+ minutes to analyze.' }
            ]
        }
    ];

    const handleAnswer = (questionId, option) => {
        const newAnswers = { ...answers, [questionId]: option };
        setAnswers(newAnswers);

        // Check if this answer disqualifies the user
        if (!option.eligible) {
            setIsEligible(false);
            setFailureReasons([option.reason]);
            setShowResult(true);
            return;
        }

        // Move to next question or show result
        if (currentQuestion < questions.length - 1) {
            setCurrentQuestion(currentQuestion + 1);
        } else {
            // All questions answered - validate
            validateAnswers(newAnswers);
        }
    };

    const validateAnswers = (finalAnswers) => {
        const reasons = [];
        let eligible = true;

        // Check all answers for eligibility
        Object.entries(finalAnswers).forEach(([questionId, answer]) => {
            if (!answer.eligible) {
                eligible = false;
                reasons.push(answer.reason);
            }
        });

        setIsEligible(eligible);
        setFailureReasons(reasons);
        setShowResult(true);
    };

    const handleRestart = () => {
        setCurrentQuestion(0);
        setAnswers({});
        setShowResult(false);
        setIsEligible(false);
        setFailureReasons([]);
    };

    const handleProceed = () => {
        const submitHandler = onSubmit || onComplete;
        if (typeof submitHandler !== 'function') {
            console.error('PreAssessmentQuestionnaire: no submit handler provided');
            return;
        }

        // Pass configuration to parent
        const config = {
            migration_type: answers.migration_type?.value,
            migrationType: answers.migration_type?.value, // backward compatibility
            requiresAI: answers.ai_analysis?.requiresAI || false,
            resourceCount: answers.resource_count?.value,
            hasWarning: answers.resource_count?.warning
        };
        submitHandler(config);
    };

    if (showResult) {
        return (
            <div className="questionnaire-container">
                <div className="questionnaire-card result-card">
                    {isEligible ? (
                        <>
                            <div className="result-icon success">✅</div>
                            <h2>You're All Set!</h2>
                            <p className="result-message">
                                This tool is perfect for your use case. You can proceed with the assessment.
                            </p>

                            {answers.resource_count?.warning && (
                                <div className="warning-box">
                                    <span className="warning-icon">⚠️</span>
                                    <p>{answers.resource_count.warning}</p>
                                </div>
                            )}

                            <div className="summary-box">
                                <h3>Assessment Summary</h3>
                                <ul>
                                    <li><strong>Migration Type:</strong> {answers.migration_type?.label}</li>
                                    <li><strong>Access Level:</strong> {answers.azure_access?.label}</li>
                                    <li><strong>Goal:</strong> {answers.assessment_goal?.label}</li>
                                    <li><strong>Reservation Data:</strong> {answers.reservation_limitation?.value === 'yes' ? 'Not Required' : 'Required'}</li>
                                    <li><strong>Environment Size:</strong> {answers.resource_count?.label}</li>
                                </ul>
                            </div>

                            <div className="result-actions">
                                <button className="btn-secondary" onClick={handleRestart}>
                                    Start Over
                                </button>
                                <button className="btn-primary" onClick={handleProceed}>
                                    Proceed to Authentication
                                </button>
                            </div>
                        </>
                    ) : (
                        <>
                            <div className="result-icon failure">❌</div>
                            <h2>This Tool May Not Be Suitable</h2>
                            <p className="result-message">
                                Based on your answers, this tool may not meet your requirements:
                            </p>

                            <div className="failure-reasons">
                                {failureReasons.map((reason, idx) => (
                                    <div key={idx} className="failure-reason">
                                        <span className="failure-icon">⚠️</span>
                                        <p>{reason}</p>
                                    </div>
                                ))}
                            </div>

                            <div className="alternative-suggestions">
                                <h3>Alternative Solutions</h3>
                                <ul>
                                    {answers.migration_type?.value === 'onprem_to_azure' && (
                                        <li>Consider using <strong>Azure Migrate</strong> for on-premises to Azure migrations</li>
                                    )}
                                    {answers.migration_type?.value === 'other_cloud_to_azure' && (
                                        <li>Use <strong>Azure Migrate</strong> or cloud-specific migration tools</li>
                                    )}
                                    {answers.migration_type?.value === 'azure_to_azure_region' && (
                                        <li>Use <strong>Azure Resource Mover</strong> for cross-region moves</li>
                                    )}
                                    {answers.assessment_goal?.value === 'cost_estimation' && (
                                        <li>Use <strong>Azure Pricing Calculator</strong> or <strong>Azure Cost Management</strong></li>
                                    )}
                                    {answers.azure_access?.value === 'no' && (
                                        <li>Request at least <strong>Reader access</strong> from your Azure administrator</li>
                                    )}
                                </ul>
                            </div>

                            <div className="result-actions">
                                <button className="btn-primary" onClick={handleRestart}>
                                    Start Over
                                </button>
                            </div>
                        </>
                    )}
                </div>
            </div>
        );
    }

    const question = questions[currentQuestion];
    const progress = ((currentQuestion + 1) / questions.length) * 100;

    return (
        <div className="questionnaire-container">
            <div className="questionnaire-card">
                <div className="progress-bar">
                    <div className="progress-fill" style={{ width: `${progress}%` }}></div>
                </div>

                <div className="question-header">
                    <span className="question-number">Question {currentQuestion + 1} of {questions.length}</span>
                    <h2>{question.question}</h2>
                </div>

                <div className="options-container">
                    {question.options.map((option, idx) => (
                        <button
                            key={idx}
                            className="option-button"
                            onClick={() => handleAnswer(question.id, option)}
                        >
                            <span className="option-label">{option.label}</span>
                            {option.warning && (
                                <span className="option-warning">⚠️ {option.warning}</span>
                            )}
                        </button>
                    ))}
                </div>

                {currentQuestion > 0 && (
                    <button
                        className="btn-back"
                        onClick={() => setCurrentQuestion(currentQuestion - 1)}
                    >
                        ← Back
                    </button>
                )}
            </div>
        </div>
    );
};

export default PreAssessmentQuestionnaire;
