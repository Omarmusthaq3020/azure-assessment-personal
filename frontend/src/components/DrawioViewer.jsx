import React, { useEffect, useRef, useState } from 'react';

// Renders a draw.io diagram from its raw XML using the official
// embed.diagrams.net iframe. Communicates via postMessage:
//   - On 'init' from the iframe → we send {action:'load', xml:...}
//   - On 'export' response → we save the resulting image
// Reference: https://www.drawio.com/doc/faq/embed-mode
//
// The embed renders in view-only mode (ui=min&saveAndExit=0). The user can
// still pan/zoom inside the iframe. Editing is not enabled — this is a
// diagram VIEWER, not an editor.

const EMBED_URL = 'https://embed.diagrams.net/?embed=1&proto=json&spin=1&ui=kennedy&libraries=0&noSaveBtn=1';

const triggerDownload = (content, filename, mime) => {
    const blob = new Blob([content], { type: `${mime};charset=utf-8` });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
};

const triggerDownloadDataUri = (dataUri, filename) => {
    const a = document.createElement('a');
    a.href = dataUri;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
};

export default function DrawioViewer({ xml }) {
    const iframeRef = useRef(null);
    const [ready, setReady] = useState(false);

    useEffect(() => {
        if (!xml) return;

        const handleMessage = (event) => {
            // Only accept messages from our iframe
            if (!iframeRef.current || event.source !== iframeRef.current.contentWindow) {
                return;
            }

            let msg;
            try {
                msg = typeof event.data === 'string' ? JSON.parse(event.data) : event.data;
            } catch {
                return; // not JSON; ignore
            }

            if (msg.event === 'init') {
                // Iframe is ready — send the diagram XML
                iframeRef.current.contentWindow.postMessage(
                    JSON.stringify({
                        action: 'load',
                        xml: xml,
                        autosave: 0,
                    }),
                    '*'
                );
                setReady(true);
            } else if (msg.event === 'export') {
                // Handle export response if we requested one
                if (msg.format === 'svg' && msg.data) {
                    triggerDownload(msg.data, 'architecture-diagram.svg', 'image/svg+xml');
                } else if (msg.format === 'png' && msg.data) {
                    // data is a base64 data URI for png
                    triggerDownloadDataUri(msg.data, 'architecture-diagram.png');
                }
            }
        };

        window.addEventListener('message', handleMessage);
        return () => window.removeEventListener('message', handleMessage);
    }, [xml]);

    const handleDownloadDrawio = () => {
        if (!xml) return;
        triggerDownload(xml, 'architecture-diagram.drawio', 'application/xml');
    };

    const handleExportSvg = () => {
        if (!iframeRef.current || !ready) return;
        iframeRef.current.contentWindow.postMessage(
            JSON.stringify({ action: 'export', format: 'svg' }),
            '*'
        );
    };

    const handleExportPng = () => {
        if (!iframeRef.current || !ready) return;
        iframeRef.current.contentWindow.postMessage(
            JSON.stringify({ action: 'export', format: 'png' }),
            '*'
        );
    };

    if (!xml) {
        return (
            <div className="w-full p-6 bg-gray-800 border border-gray-700 rounded-lg text-gray-400">
                No diagram XML to display.
            </div>
        );
    }

    return (
        <div className="w-full flex-col">
            <div className="flex gap-2 mb-4 justify-end">
                <button
                    onClick={handleDownloadDrawio}
                    className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs rounded-md shadow flex items-center transition"
                >
                    Download .drawio
                </button>
                <button
                    onClick={handleExportSvg}
                    disabled={!ready}
                    className={`px-3 py-1.5 text-white text-xs rounded-md shadow flex items-center transition ${
                        ready ? 'bg-gray-700 hover:bg-gray-600' : 'bg-gray-900 cursor-not-allowed'
                    }`}
                >
                    Export SVG
                </button>
                <button
                    onClick={handleExportPng}
                    disabled={!ready}
                    className={`px-3 py-1.5 text-white text-xs rounded-md shadow flex items-center transition ${
                        ready ? 'bg-gray-700 hover:bg-gray-600' : 'bg-gray-900 cursor-not-allowed'
                    }`}
                >
                    Export PNG
                </button>
            </div>

            <iframe
                ref={iframeRef}
                title="Architecture diagram"
                src={EMBED_URL}
                className="w-full bg-white border border-gray-300 rounded-lg"
                style={{ height: '700px', border: 0 }}
                sandbox="allow-scripts allow-same-origin allow-downloads allow-popups"
            />
        </div>
    );
}
