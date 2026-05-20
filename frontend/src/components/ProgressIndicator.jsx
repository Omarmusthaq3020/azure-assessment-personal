import { useState, useEffect } from 'react';

const migrationTips = [
    "💡 Tip: Standard SKU Public IPs must be detached before moving across subscriptions",
    "⚡ Did you know? VNet peerings must be deleted before subscription moves",
    "🔍 The AI analyzes resources against official Microsoft documentation",
    "🚀 Resources in the same resource group can be moved together",
    "⚠️ Some resources like Azure AD Domain Services cannot be moved",
    "📦 Managed disks and NICs move automatically with their VMs",
    "🔐 Key Vault secrets and certificates are preserved during moves",
    "🌐 DNS records may need to be updated after moving resources",
    "⏱️ Large environments may take 10-15 minutes to analyze",
    "✅ This tool checks 100+ Azure resource types for move compatibility",
    "🔄 Cross-subscription moves require proper RBAC permissions",
    "📊 The dashboard provides detailed insights into blocking factors",
    "🎯 Review Required items need manual verification before moving",
    "🛡️ Network Security Groups can be moved with their associated resources",
    "💾 Storage accounts must have no active leases during the move"
];

export default function ProgressIndicator({ progress }) {
    const [currentTipIndex, setCurrentTipIndex] = useState(0);

    useEffect(() => {
        const interval = setInterval(() => {
            setCurrentTipIndex((prev) => (prev + 1) % migrationTips.length);
        }, 8000); // Rotate every 8 seconds

        return () => clearInterval(interval);
    }, []);

    const percentage = progress?.percentage || 0;
    const step = progress?.step || "Initializing";
    const message = progress?.message || "Starting assessment...";

    return (
        <div className="space-y-6">
            {/* Progress Header */}
            <div className="text-center">
                <div className="animate-spin rounded-full h-16 w-16 border-t-2 border-b-2 border-blue-500 mx-auto mb-4"></div>
                <h3 className="text-xl font-semibold text-blue-400 mb-2">🔄 Running Assessment...</h3>
            </div>

            {/* Progress Bar */}
            <div className="space-y-2">
                <div className="flex justify-between text-sm">
                    <span className="text-gray-300 font-medium">{step}</span>
                    <span className="text-blue-400 font-bold">{percentage}%</span>
                </div>
                <div className="w-full bg-gray-700 rounded-full h-3 overflow-hidden">
                    <div
                        className="bg-gradient-to-r from-blue-500 to-purple-500 h-3 rounded-full transition-all duration-500 ease-out"
                        style={{ width: `${percentage}%` }}
                    ></div>
                </div>
                {message && (
                    <p className="text-xs text-gray-400 italic">{message}</p>
                )}
            </div>

            {/* Rotating Tips */}
            <div className="bg-gray-700/50 border border-gray-600 rounded-lg p-4 min-h-[80px] flex items-center justify-center">
                <p
                    key={currentTipIndex}
                    className="text-sm text-gray-300 text-center animate-fade-in"
                >
                    {migrationTips[currentTipIndex]}
                </p>
            </div>

            {/* Additional Info */}
            <p className="text-xs text-gray-500 text-center">
                This may take a few minutes depending on your environment size.
            </p>
        </div>
    );
}
