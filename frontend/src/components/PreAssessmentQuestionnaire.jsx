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
        Object.values(finalAnswers).forEach((answer) => {
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
            azure_access: answers.azure_access?.value,
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

                            <div className="summary-box">
                                <h3>Assessment Summary</h3>
                                <ul>
                                    <li><strong>Migration Type:</strong> {answers.migration_type?.label}</li>
                                    <li><strong>Access Level:</strong> {answers.azure_access?.label}</li>
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
                                    {answers.migration_type?.value === 'other_cloud_to_azure' && (
                                        <li>Use <strong>Azure Migrate</strong> or cloud-specific migration tools</li>
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
