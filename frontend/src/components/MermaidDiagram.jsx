import React, { useEffect, useRef, useState } from 'react';
import mermaid from 'mermaid';

export default function MermaidDiagram({ code, drawioXml = '' }) {
    const containerRef = useRef(null);
    const [svgContent, setSvgContent] = useState('');
    const [error, setError] = useState('');

    useEffect(() => {
        mermaid.initialize({
            startOnLoad: false,
            theme: 'neutral',
            securityLevel: 'loose',
            logLevel: 5,
            flowchart: {
                useMaxWidth: true,
                htmlLabels: true,
                curve: 'basis'
            }
        });

        const renderDiagram = async () => {
            if (!code || !containerRef.current) return;

            try {
                const id = `mermaid-${Math.random().toString(36).substr(2, 9)}`;
                setError('');
                const { svg } = await mermaid.render(id, code);
                setSvgContent(svg);
            } catch (err) {
                console.error('Mermaid rendering failed:', err);
                setError(err.message || 'Failed to render diagram');
                setSvgContent('');
            }
        };

        renderDiagram();
    }, [code]);

    const handleDownloadSVG = () => {
        if (!svgContent) return;
        const blob = new Blob([svgContent], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'architecture-diagram.svg';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    };

    const handleDownloadPNG = () => {
        if (!containerRef.current) return;
        const svgElement = containerRef.current.querySelector('svg');
        if (!svgElement) return;

        const viewBox = svgElement.getAttribute('viewBox')?.split(' ');
        let width = 1000;
        let height = 1000;
        if (viewBox && viewBox.length === 4) {
            width = parseFloat(viewBox[2]);
            height = parseFloat(viewBox[3]);
        } else {
            const bbox = svgElement.getBoundingClientRect();
            width = bbox.width || 1000;
            height = bbox.height || 1000;
        }

        const clonedSvg = svgElement.cloneNode(true);
        clonedSvg.style.backgroundColor = '#ffffff';
        const svgData = new XMLSerializer().serializeToString(clonedSvg);

        const canvas = document.createElement('canvas');
        const scale = 2;
        canvas.width = width * scale;
        canvas.height = height * scale;

        const ctx = canvas.getContext('2d');
        const img = new Image();
        const svgBlob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(svgBlob);

        img.onload = () => {
            ctx.scale(scale, scale);
            ctx.drawImage(img, 0, 0, width, height);
            URL.revokeObjectURL(url);
            const pngUrl = canvas.toDataURL('image/png');
            const a = document.createElement('a');
            a.href = pngUrl;
            a.download = 'architecture-diagram.png';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        };
        img.src = url;
    };

    const handleCopyDrawIo = () => {
        navigator.clipboard.writeText(code);
        alert('Source code copied. In draw.io: Arrange > Insert > Advanced > Mermaid');
    };

    const handleDownloadDrawio = () => {
        if (!drawioXml) return;
        const blob = new Blob([drawioXml], { type: 'application/xml;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'architecture-diagram.drawio';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    };

    return (
        <div className="w-full flex-col">
            {!error && svgContent && (
                <div className="flex gap-2 mb-4 justify-end">
                    <button
                        onClick={handleCopyDrawIo}
                        className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs rounded-md shadow flex items-center transition"
                    >
                        Copy for Draw.io
                    </button>
                    <button
                        onClick={handleDownloadDrawio}
                        disabled={!drawioXml}
                        className={`px-3 py-1.5 text-white text-xs rounded-md shadow flex items-center transition ${
                            drawioXml ? 'bg-emerald-600 hover:bg-emerald-500' : 'bg-emerald-900 cursor-not-allowed'
                        }`}
                    >
                        Download .drawio
                    </button>
                    <button
                        onClick={handleDownloadSVG}
                        className="px-3 py-1.5 bg-gray-700 hover:bg-gray-600 text-white text-xs rounded-md shadow flex items-center transition"
                    >
                        Export SVG
                    </button>
                    <button
                        onClick={handleDownloadPNG}
                        className="px-3 py-1.5 bg-gray-700 hover:bg-gray-600 text-white text-xs rounded-md shadow flex items-center transition"
                    >
                        Export PNG
                    </button>
                </div>
            )}

            {error ? (
                <div className="p-4 bg-red-900/30 border border-red-700 text-red-400 rounded-lg whitespace-pre-wrap font-mono text-xs">
                    Diagram Syntax Error:
                    {error}
                </div>
            ) : (
                <div
                    ref={containerRef}
                    className="w-full overflow-auto bg-white border border-gray-300 rounded-lg p-6 flex justify-center items-center min-h-[300px]"
                    dangerouslySetInnerHTML={{ __html: svgContent }}
                />
            )}
        </div>
    );
}
