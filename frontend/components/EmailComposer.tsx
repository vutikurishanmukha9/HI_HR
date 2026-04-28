
import React, { useState, useRef, useEffect, DragEvent } from 'react';
import { EmailTemplate } from '../types';

interface EmailComposerProps {
    initialTemplate: EmailTemplate;
    onCompose: (template: EmailTemplate) => void;
    onBack: () => void;
    recipients?: any[];
}

const personalizationTags = ['{fullName}', '{companyName}', '{jobTitle}'];
const MAX_ATTACHMENT_SIZE = 10 * 1024 * 1024;
const MAX_SUBJECT_LENGTH = 200;

const EmailComposer: React.FC<EmailComposerProps> = ({ initialTemplate, onCompose, onBack }) => {
    const [subject, setSubject] = useState(initialTemplate.subject);
    const [body, setBody] = useState(initialTemplate.body);
    const [attachments, setAttachments] = useState<File[]>(initialTemplate.attachments || []);
    const editorRef = useRef<HTMLDivElement>(null);
    const [showPlaceholder, setShowPlaceholder] = useState(false);
    const [isDraggingAttachment, setIsDraggingAttachment] = useState(false);
    const [attachmentError, setAttachmentError] = useState<string | null>(null);
    const [copiedTag, setCopiedTag] = useState<string | null>(null);

    // Robust clipboard copy with fallback
    const copyToClipboard = async (text: string) => {
        try {
            // Modern API (requires HTTPS or localhost)
            if (navigator.clipboard && window.isSecureContext) {
                await navigator.clipboard.writeText(text);
            } else {
                // Fallback for older browsers or non-secure contexts
                const textArea = document.createElement('textarea');
                textArea.value = text;
                textArea.style.position = 'fixed';
                textArea.style.left = '-999999px';
                textArea.style.top = '-999999px';
                document.body.appendChild(textArea);
                textArea.focus();
                textArea.select();
                document.execCommand('copy');
                document.body.removeChild(textArea);
            }
            // Show feedback
            setCopiedTag(text);
            setTimeout(() => setCopiedTag(null), 1500);
        } catch (err) {
            console.error('Copy failed:', err);
            alert('Copy failed. Please manually select and copy: ' + text);
        }
    };

    useEffect(() => {
        if (editorRef.current && body !== editorRef.current.innerHTML) {
            editorRef.current.innerHTML = body;
        }
        if (editorRef.current) {
            setShowPlaceholder(!editorRef.current.innerText.trim() && !editorRef.current.innerHTML.includes('<img'));
        }
    }, [body]);

    const processFiles = (files: FileList) => {
        setAttachmentError(null);
        if (files && files.length > 0) {
            const validFiles: File[] = [];
            const invalidFiles: string[] = [];
            Array.from(files).forEach(file => {
                if (file.size > MAX_ATTACHMENT_SIZE) {
                    invalidFiles.push(`${file.name} (${(file.size / 1024 / 1024).toFixed(2)}MB)`);
                } else {
                    validFiles.push(file);
                }
            });
            if (invalidFiles.length > 0) {
                setAttachmentError(`Files exceed 10MB limit: ${invalidFiles.join(', ')}`);
            }
            if (validFiles.length > 0) {
                setAttachments(prev => [...prev, ...validFiles]);
            }
        }
    };

    const handleAttachmentChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        if (e.target.files) processFiles(e.target.files);
        e.target.value = '';
    };

    const removeAttachment = (fileName: string) => {
        setAttachments(prev => prev.filter(file => file.name !== fileName));
        setAttachmentError(null);
    };

    const handleBodyChange = (e: React.FormEvent<HTMLDivElement>) => {
        const editor = e.currentTarget;
        if (body !== editor.innerHTML) setBody(editor.innerHTML);
    };

    const handleFormat = (command: string, value?: string) => {
        document.execCommand(command, false, value);
        editorRef.current?.focus();
        if (editorRef.current) setBody(editorRef.current.innerHTML);
    };

    const handleLink = () => {
        const url = prompt('Enter the URL:');
        if (url) handleFormat('createLink', url);
    };

    const handleKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
        if (e.ctrlKey) {
            if (e.key === 'z') { e.preventDefault(); handleFormat('undo'); }
            else if (e.key === 'y') { e.preventDefault(); handleFormat('redo'); }
        }
    };

    const handleSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        if (subject.trim() && body.trim()) onCompose({ subject, body, attachments });
    };

    const handleAttachmentDragEnter = (e: DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); setIsDraggingAttachment(true); };
    const handleAttachmentDragLeave = (e: DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); setIsDraggingAttachment(false); };
    const handleAttachmentDragOver = (e: DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); };
    const handleAttachmentDrop = (e: DragEvent<HTMLDivElement>) => {
        e.preventDefault(); e.stopPropagation(); setIsDraggingAttachment(false);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            processFiles(e.dataTransfer.files);
            e.dataTransfer.clearData();
        }
    };

    const ToolbarButton: React.FC<{ onClick: () => void; children: React.ReactNode, title: string }> = ({ onClick, children, title }) => (
        <button type="button" title={title} onClick={onClick} onMouseDown={(e) => e.preventDefault()}
            className="p-2 rounded-lg transition-colors hover:bg-white/10">
            {children}
        </button>
    );

    return (
        <div className="max-w-4xl mx-auto fade-in">
            {/* Header */}
            <div className="text-center mb-8">
                <div className="inline-flex items-center justify-center w-14 h-14 rounded-2xl mb-4"
                    style={{ background: 'linear-gradient(135deg, rgba(244, 63, 94, 0.15), rgba(249, 115, 22, 0.15))', border: '1px solid rgba(244, 63, 94, 0.25)' }}>
                    <svg className="w-7 h-7" style={{ color: '#fb7185' }} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                    </svg>
                </div>
                <h2 className="text-2xl font-bold mb-2" style={{ color: '#f1f5f9' }}>Compose Your Email</h2>
                <p style={{ color: '#94a3b8', fontSize: '0.9375rem' }}>Craft your message with personalization tags</p>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                {/* Main Editor */}
                <div className="lg:col-span-2 space-y-6">
                    <form onSubmit={handleSubmit} className="space-y-6">
                        {/* Subject */}
                        <div>
                            <label className="block text-sm font-medium text-gray-300 mb-2">Subject Line</label>
                            <input
                                type="text"
                                value={subject}
                                onChange={(e) => setSubject(e.target.value)}
                                placeholder="An opportunity at {companyName}"
                                required
                                maxLength={MAX_SUBJECT_LENGTH}
                            />
                            <p className="mt-2 text-xs text-gray-500 text-right">{subject.length}/{MAX_SUBJECT_LENGTH}</p>
                        </div>

                        {/* Editor */}
                        <div>
                            <label className="block text-sm font-medium text-gray-300 mb-2">Email Body</label>
                            <div className="rounded-xl overflow-hidden" style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.1)' }}>
                                {/* Toolbar */}
                                <div className="flex items-center gap-1 p-3 border-b" style={{ borderColor: 'rgba(255,255,255,0.1)', background: 'rgba(255,255,255,0.02)' }}>
                                    <ToolbarButton title="Undo" onClick={() => handleFormat('undo')}>
                                        <svg className="w-4 h-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M11 15l-3-3m0 0l3-3m-3 3h8a5 5 0 000-10H9"></path></svg>
                                    </ToolbarButton>
                                    <ToolbarButton title="Redo" onClick={() => handleFormat('redo')}>
                                        <svg className="w-4 h-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 15l3-3m0 0l-3-3m3 3H5a5 5 0 000 10h6"></path></svg>
                                    </ToolbarButton>
                                    <div className="w-px h-4 mx-1" style={{ background: 'rgba(255,255,255,0.1)' }} />
                                    <ToolbarButton title="Bold" onClick={() => handleFormat('bold')}>
                                        <span className="font-bold text-gray-400">B</span>
                                    </ToolbarButton>
                                    <ToolbarButton title="Italic" onClick={() => handleFormat('italic')}>
                                        <span className="italic text-gray-400">I</span>
                                    </ToolbarButton>
                                    <ToolbarButton title="List" onClick={() => handleFormat('insertUnorderedList')}>
                                        <svg className="w-4 h-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 6h16M4 12h16M4 18h16"></path></svg>
                                    </ToolbarButton>
                                    <ToolbarButton title="Link" onClick={handleLink}>
                                        <svg className="w-4 h-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101"></path></svg>
                                    </ToolbarButton>
                                </div>

                                {/* Content Editable */}
                                <div className="relative">
                                    {showPlaceholder && (
                                        <div className="absolute top-4 left-4 text-gray-600 pointer-events-none whitespace-pre-wrap">
                                            {`Hi {fullName},\n\nI came across your profile as a {jobTitle} at {companyName}...`}
                                        </div>
                                    )}
                                    <div
                                        ref={editorRef}
                                        contentEditable={true}
                                        onInput={handleBodyChange}
                                        onKeyDown={handleKeyDown}
                                        className="min-h-[250px] p-4 text-white focus:outline-none prose-invert"
                                        style={{ background: 'transparent' }}
                                    />
                                </div>
                            </div>
                        </div>
                    </form>

                    {/* Attachments */}
                    <div>
                        <label className="block text-sm font-medium text-gray-300 mb-2">Attachments</label>
                        {attachmentError && (
                            <div className="mb-3 p-3 rounded-lg" style={{ background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.3)' }}>
                                <p className="text-red-400 text-sm">{attachmentError}</p>
                            </div>
                        )}
                        <div
                            className={`p-6 rounded-xl text-center cursor-pointer transition-all duration-300`}
                            style={{
                                background: isDraggingAttachment ? 'rgba(102, 126, 234, 0.1)' : 'rgba(255,255,255,0.02)',
                                border: isDraggingAttachment ? '2px dashed #667eea' : '2px dashed rgba(255,255,255,0.1)',
                            }}
                            onDragEnter={handleAttachmentDragEnter}
                            onDragLeave={handleAttachmentDragLeave}
                            onDragOver={handleAttachmentDragOver}
                            onDrop={handleAttachmentDrop}
                            onClick={() => document.getElementById('attachment-upload')?.click()}
                        >
                            <svg className="w-8 h-8 text-gray-500 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15.172 7l-6.586 6.586a2 2 0 102.828 2.828l6.414-6.586a4 4 0 00-5.656-5.656l-6.415 6.585a6 6 0 108.486 8.486L20.5 13" />
                            </svg>
                            <p className="text-gray-400 text-sm">Drop files or click to browse</p>
                            <input id="attachment-upload" type="file" className="sr-only" multiple onChange={handleAttachmentChange} />
                        </div>

                        {attachments.length > 0 && (
                            <div className="mt-3 space-y-2">
                                {attachments.map((file, index) => (
                                    <div key={index} className="flex items-center justify-between p-3 rounded-lg"
                                        style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.1)' }}>
                                        <div className="flex items-center gap-3">
                                            <svg className="w-5 h-5 text-purple-400" fill="currentColor" viewBox="0 0 20 20">
                                                <path fillRule="evenodd" d="M8 4a3 3 0 00-3 3v4a3 3 0 006 0V7a1 1 0 112 0v4a5 5 0 01-10 0V7a5 5 0 0110 0v4a1 1 0 11-2 0V7a3 3 0 00-3-3z" clipRule="evenodd" />
                                            </svg>
                                            <span className="text-sm text-white truncate max-w-[200px]">{file.name}</span>
                                        </div>
                                        <button onClick={() => removeAttachment(file.name)} className="text-red-400 hover:text-red-300 text-sm">
                                            Remove
                                        </button>
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                </div>

                {/* Sidebar */}
                <aside>
                    <div className="rounded-xl p-5" style={{ background: 'rgba(244, 63, 94, 0.06)', border: '1px solid rgba(244, 63, 94, 0.15)' }}>
                        <h4 className="font-semibold flex items-center gap-2 mb-2" style={{ color: '#f1f5f9' }}>
                            <svg className="w-5 h-5" style={{ color: '#fb7185' }} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 7h.01M7 3h5c.512 0 1.024.195 1.414.586l7 7a2 2 0 010 2.828l-7 7a2 2 0 01-2.828 0l-7-7A1.994 1.994 0 013 12V7a4 4 0 014-4z" />
                            </svg>
                            Personalization Tags
                        </h4>
                        <p className="text-sm mb-4" style={{ color: '#94a3b8' }}>Click to copy, then paste into your email</p>
                        <div className="space-y-2">
                            {personalizationTags.map(tag => (
                                <button
                                    key={tag}
                                    onClick={() => copyToClipboard(tag)}
                                    className="w-full text-left px-4 py-3 text-sm rounded-lg transition-all duration-300 hover:scale-[1.02] flex items-center justify-between"
                                    style={{
                                        background: copiedTag === tag ? 'rgba(20, 184, 166, 0.15)' : 'rgba(148, 163, 184, 0.06)',
                                        border: copiedTag === tag ? '1px solid rgba(20, 184, 166, 0.4)' : '1px solid rgba(148, 163, 184, 0.12)',
                                        color: copiedTag === tag ? '#14b8a6' : '#fb7185'
                                    }}
                                >
                                    <span>{tag}</span>
                                    {copiedTag === tag && (
                                        <span className="text-xs flex items-center gap-1" style={{ color: '#14b8a6' }}>
                                            <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
                                                <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                                            </svg>
                                            Copied!
                                        </span>
                                    )}
                                </button>
                            ))}
                        </div>
                    </div>
                </aside>
            </div>

            {/* Action Buttons */}
            <div className="mt-8 flex justify-between">
                <button
                    type="button"
                    onClick={onBack}
                    className="px-6 py-3 rounded-xl font-medium text-sm transition-all duration-300"
                    style={{ border: '1px solid rgba(148, 163, 184, 0.15)', color: '#94a3b8' }}
                >
                    ← Back
                </button>
                <button
                    type="submit"
                    onClick={handleSubmit}
                    disabled={!subject.trim() || !body.trim()}
                    className="px-8 py-3 rounded-xl font-semibold text-white text-sm transition-all duration-300 disabled:opacity-40 disabled:cursor-not-allowed hover:scale-[1.02] active:scale-[0.98]"
                    style={{
                        background: 'linear-gradient(135deg, #f43f5e 0%, #f97316 100%)',
                        boxShadow: (!subject.trim() || !body.trim()) ? 'none' : '0 4px 20px rgba(244, 63, 94, 0.35)'
                    }}
                >
                    Continue →
                </button>
            </div>
        </div>
    );
};

export default EmailComposer;