import React, { useState, useRef, useEffect, useMemo, DragEvent } from 'react';
import { EmailTemplate } from '../types';
import TemplateModal from './TemplateModal';
import { analyzeSpamScore } from '../utils/spamChecker';

interface EmailComposerProps {
    initialTemplate: EmailTemplate;
    onCompose: (template: EmailTemplate) => void;
    onBack: () => void;
    recipients?: any[];
    availableTags?: string[];
}
const MAX_ATTACHMENT_SIZE = 10 * 1024 * 1024;
const MAX_SUBJECT_LENGTH = 200;

const EmailComposer: React.FC<EmailComposerProps> = ({ initialTemplate, onCompose, onBack, availableTags = [] }) => {
    const [subject, setSubject] = useState(initialTemplate?.subject || '');
    const [body, setBody] = useState(initialTemplate?.body || '');
    const [attachments, setAttachments] = useState<File[]>(initialTemplate?.attachments || []);
    const editorRef = useRef<HTMLDivElement>(null);
    const [showPlaceholder, setShowPlaceholder] = useState(false);
    const [isDraggingAttachment, setIsDraggingAttachment] = useState(false);
    const [attachmentError, setAttachmentError] = useState<string | null>(null);
    const [copiedTag, setCopiedTag] = useState<string | null>(null);
    const [isTemplateModalOpen, setIsTemplateModalOpen] = useState(false);

    // Compute live deliverability spam score
    const spamReport = useMemo(() => analyzeSpamScore(subject, body), [subject, body]);

    // Robust clipboard copy with fallback
    const copyToClipboard = async (text: string) => {
        try {
            if (navigator.clipboard && window.isSecureContext) {
                await navigator.clipboard.writeText(text);
            } else {
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
            setCopiedTag(text);
            setTimeout(() => setCopiedTag(null), 1500);
        } catch (err) {
            console.error('Copy failed:', err);
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
        <button
            type="button"
            title={title}
            onClick={onClick}
            onMouseDown={(e) => e.preventDefault()}
            className="p-1.5 rounded-md text-ink-secondary hover:text-ink hover:bg-neutral-100 transition-colors"
        >
            {children}
        </button>
    );

    const scoreColors = useMemo(() => {
        if (spamReport.score >= 90) return { text: 'text-emerald-700', bg: 'bg-emerald-50', border: 'border-emerald-200', bar: 'bg-emerald-500' };
        if (spamReport.score >= 75) return { text: 'text-amber-700', bg: 'bg-amber-50', border: 'border-amber-200', bar: 'bg-amber-500' };
        return { text: 'text-red-700', bg: 'bg-red-50', border: 'border-red-200', bar: 'bg-red-500' };
    }, [spamReport.score]);

    return (
        <div className="max-w-4xl mx-auto">
            {/* Header */}
            <div className="text-center mb-6">
                <div className="inline-flex items-center justify-center w-12 h-12 rounded-xl bg-accent-light border border-neutral-200 text-accent mb-3 shadow-xs">
                    <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                    </svg>
                </div>
                <h2 className="text-xl font-bold text-ink">Compose Your Email</h2>
                <p className="text-sm text-ink-muted mt-1">Craft personalized cold email copy using variables</p>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                {/* Main Editor Column */}
                <div className="lg:col-span-2 space-y-5">
                    {/* Top Action Bar */}
                    <div className="flex flex-wrap justify-between items-center gap-2 p-3 bg-neutral-50 border border-border rounded-xl">
                        <button
                            type="button"
                            onClick={() => setIsTemplateModalOpen(true)}
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-card hover:bg-neutral-100 text-ink border border-neutral-200 shadow-xs transition-colors"
                        >
                            <svg className="w-4 h-4 text-accent" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" />
                            </svg>
                            Templates Library
                        </button>

                        <div className="flex items-center gap-2">
                            <span className="text-[11px] text-ink-muted font-medium">Deliverability Score:</span>
                            <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold border ${scoreColors.bg} ${scoreColors.text} ${scoreColors.border}`}>
                                <span className={`w-2 h-2 rounded-full ${scoreColors.bar}`} />
                                {spamReport.score}/100 - {spamReport.rating}
                            </span>
                        </div>
                    </div>

                    <form onSubmit={handleSubmit} className="space-y-4">
                        {/* Subject */}
                        <div className="space-y-1.5">
                            <div className="flex justify-between items-center">
                                <label className="block text-xs font-semibold uppercase tracking-wider text-ink-secondary">
                                    Subject Line
                                </label>
                                <span className="text-[11px] text-ink-muted tabular-nums">
                                    {(subject || '').length}/{MAX_SUBJECT_LENGTH}
                                </span>
                            </div>
                            <input
                                type="text"
                                value={subject}
                                onChange={(e) => setSubject(e.target.value)}
                                placeholder="Opportunity at {companyName}"
                                className="input"
                                required
                                maxLength={MAX_SUBJECT_LENGTH}
                            />
                        </div>

                        {/* Editor Box */}
                        <div className="space-y-1.5">
                            <label className="block text-xs font-semibold uppercase tracking-wider text-ink-secondary">
                                Email Body
                            </label>
                            <div className="border border-border rounded-xl overflow-hidden shadow-xs bg-card focus-within:border-accent focus-within:ring-2 focus-within:ring-neutral-200 transition-all">
                                {/* Toolbar */}
                                <div className="flex items-center gap-1 px-3 py-2 border-b border-border bg-neutral-50">
                                    <ToolbarButton title="Undo (Ctrl+Z)" onClick={() => handleFormat('undo')}>
                                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M11 15l-3-3m0 0l3-3m-3 3h8a5 5 0 000-10H9"></path></svg>
                                    </ToolbarButton>
                                    <ToolbarButton title="Redo (Ctrl+Y)" onClick={() => handleFormat('redo')}>
                                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 15l3-3m0 0l-3-3m3 3H5a5 5 0 000 10h6"></path></svg>
                                    </ToolbarButton>
                                    <div className="w-px h-4 mx-1 bg-border" />
                                    <ToolbarButton title="Bold" onClick={() => handleFormat('bold')}>
                                        <span className="font-bold text-xs px-0.5">B</span>
                                    </ToolbarButton>
                                    <ToolbarButton title="Italic" onClick={() => handleFormat('italic')}>
                                        <span className="italic font-serif text-xs px-0.5">I</span>
                                    </ToolbarButton>
                                    <ToolbarButton title="Bullet List" onClick={() => handleFormat('insertUnorderedList')}>
                                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 6h16M4 12h16M4 18h16"></path></svg>
                                    </ToolbarButton>
                                    <ToolbarButton title="Insert Link" onClick={handleLink}>
                                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101"></path></svg>
                                    </ToolbarButton>
                                </div>

                                {/* Content Editable Body */}
                                <div className="relative">
                                    {showPlaceholder && (
                                        <div className="absolute top-4 left-4 pointer-events-none text-ink-muted text-sm leading-relaxed whitespace-pre-wrap">
                                            {`Hi {name},\n\nI came across your profile and was impressed by your work at {company}...`}
                                        </div>
                                    )}
                                    <div
                                        ref={editorRef}
                                        contentEditable={true}
                                        onInput={handleBodyChange}
                                        onKeyDown={handleKeyDown}
                                        className="min-h-[260px] p-4 focus:outline-none text-sm leading-relaxed text-ink"
                                    />
                                </div>
                            </div>
                        </div>
                    </form>

                    {/* Attachments Section */}
                    <div className="space-y-2">
                        <label className="block text-xs font-semibold uppercase tracking-wider text-ink-secondary">
                            Attachments
                        </label>
                        {attachmentError && (
                            <div className="p-3 rounded-lg bg-danger-light border border-red-200 text-danger text-xs font-medium animate-fadeIn">
                                {attachmentError}
                            </div>
                        )}
                        <div
                            className={`p-5 rounded-xl text-center cursor-pointer border-2 border-dashed transition-all ${
                                isDraggingAttachment
                                    ? 'border-accent bg-neutral-100'
                                    : 'border-neutral-300 hover:border-accent bg-neutral-50'
                            }`}
                            onDragEnter={handleAttachmentDragEnter}
                            onDragLeave={handleAttachmentDragLeave}
                            onDragOver={handleAttachmentDragOver}
                            onDrop={handleAttachmentDrop}
                            onClick={() => document.getElementById('attachment-upload')?.click()}
                        >
                            <svg className="w-6 h-6 mx-auto mb-1 text-ink-muted" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.8} d="M15.172 7l-6.586 6.586a2 2 0 102.828 2.828l6.414-6.586a4 4 0 00-5.656-5.656l-6.415 6.585a6 6 0 108.486 8.486L20.5 13" />
                            </svg>
                            <p className="text-xs font-medium text-ink">Drop attachment files or browse</p>
                            <p className="text-[11px] text-ink-muted mt-0.5">Resume, portfolio, PDF up to 10MB</p>
                            <input id="attachment-upload" type="file" className="sr-only" multiple onChange={handleAttachmentChange} />
                        </div>

                        {attachments.length > 0 && (
                            <div className="space-y-1.5 mt-2">
                                {attachments.map((file, index) => (
                                    <div
                                        key={index}
                                        className="flex items-center justify-between p-2.5 rounded-lg bg-neutral-50 border border-border text-xs"
                                    >
                                        <div className="flex items-center gap-2 truncate pr-2">
                                            <svg className="w-4 h-4 text-accent flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                                                <path fillRule="evenodd" d="M8 4a3 3 0 00-3 3v4a3 3 0 006 0V7a1 1 0 112 0v4a5 5 0 01-10 0V7a5 5 0 0110 0v4a1 1 0 11-2 0V7a3 3 0 00-3-3z" clipRule="evenodd" />
                                            </svg>
                                            <span className="font-medium text-ink truncate">{file.name}</span>
                                            <span className="text-ink-muted text-[10px]">({(file.size / 1024).toFixed(0)} KB)</span>
                                        </div>
                                        <button
                                            type="button"
                                            onClick={() => removeAttachment(file.name)}
                                            className="text-danger hover:text-red-700 font-medium text-[11px]"
                                        >
                                            Remove
                                        </button>
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                </div>

                {/* Sidebar: Dynamic Tags & Deliverability Insights */}
                <aside className="space-y-5">
                    {/* Personalization Tags */}
                    <div className="surface-card p-4 sm:p-5 border border-border shadow-xs">
                        <div className="flex items-center gap-2 mb-1.5">
                            <div className="w-7 h-7 rounded-lg bg-neutral-100 border border-neutral-200 flex items-center justify-center text-accent">
                                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 7h.01M7 3h5c.512 0 1.024.195 1.414.586l7 7a2 2 0 010 2.828l-7 7a2 2 0 01-2.828 0l-7-7A1.994 1.994 0 013 12V7a4 4 0 014-4z" />
                                </svg>
                            </div>
                            <h3 className="font-semibold text-sm text-ink">Personalization Tags</h3>
                        </div>
                        <p className="text-xs text-ink-muted mb-3">Click any variable to copy it to your clipboard</p>

                        {(availableTags || []).length === 0 ? (
                            <div className="text-xs p-3.5 rounded-lg text-center bg-neutral-50 border border-border text-ink-muted">
                                Upload a recipient list to automatically detect available columns.
                            </div>
                        ) : (
                            <div className="space-y-1.5 max-h-[220px] overflow-y-auto pr-1">
                                {availableTags.map(tag => {
                                    const formattedTag = `{${tag}}`;
                                    const isCopied = copiedTag === formattedTag;

                                    return (
                                        <button
                                            key={tag}
                                            type="button"
                                            onClick={() => copyToClipboard(formattedTag)}
                                            className={`w-full text-left px-3 py-2 rounded-lg text-xs font-mono font-medium transition-all flex items-center justify-between border gap-2 ${
                                                isCopied
                                                    ? 'bg-emerald-50 border-emerald-300 text-emerald-700'
                                                    : 'bg-neutral-50 hover:bg-neutral-100 border-border hover:border-neutral-300 text-ink'
                                            }`}
                                        >
                                            <span className="truncate min-w-0 flex-1" title={formattedTag}>
                                                {formattedTag}
                                            </span>
                                            {isCopied ? (
                                                <span className="text-[10px] text-emerald-700 font-sans font-semibold flex items-center gap-1 flex-shrink-0">
                                                    <svg className="w-3 h-3 text-emerald-600 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M5 13l4 4L19 7" /></svg>
                                                    <span>Copied</span>
                                                </span>
                                            ) : (
                                                <span className="text-[10px] text-ink-muted font-sans flex-shrink-0">
                                                    Click to copy
                                                </span>
                                            )}
                                        </button>
                                    );
                                })}
                            </div>
                        )}
                    </div>

                    {/* Deliverability & Spam Insights */}
                    <div className="surface-card p-4 sm:p-5 border border-border shadow-xs">
                        <div className="flex items-center justify-between mb-2">
                            <div className="flex items-center gap-2">
                                <div className="w-7 h-7 rounded-lg bg-neutral-100 border border-neutral-200 flex items-center justify-center text-accent">
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
                                    </svg>
                                </div>
                                <h3 className="font-semibold text-sm text-ink">Deliverability Insights</h3>
                            </div>
                            <span className={`text-xs font-bold px-2 py-0.5 rounded-full border ${scoreColors.bg} ${scoreColors.text} ${scoreColors.border}`}>
                                {spamReport.score}/100
                            </span>
                        </div>

                        {/* Meter Bar */}
                        <div className="w-full bg-neutral-100 rounded-full h-2 my-3 overflow-hidden border border-neutral-200">
                            <div
                                className={`h-full transition-all duration-300 ${scoreColors.bar}`}
                                style={{ width: `${spamReport.score}%` }}
                            />
                        </div>

                        {/* Issues list */}
                        {spamReport.issues.length === 0 ? (
                            <div className="p-3 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs flex items-center gap-2">
                                <svg className="w-4 h-4 text-emerald-600 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                                </svg>
                                <span>High inbox placement potential. No spam triggers detected.</span>
                            </div>
                        ) : (
                            <div className="space-y-2 max-h-[220px] overflow-y-auto pr-1">
                                {spamReport.issues.map((issue, idx) => (
                                    <div
                                        key={idx}
                                        className={`p-2.5 rounded-lg border text-xs leading-relaxed ${
                                            issue.type === 'high'
                                                ? 'bg-red-50 border-red-200 text-red-800'
                                                : issue.type === 'medium'
                                                ? 'bg-amber-50 border-amber-200 text-amber-800'
                                                : 'bg-neutral-50 border-neutral-200 text-ink-secondary'
                                        }`}
                                    >
                                        <div className="font-semibold text-[10px] uppercase tracking-wider mb-0.5">
                                            {issue.type === 'high' ? 'High Risk' : issue.type === 'medium' ? 'Notice' : 'Suggestion'}
                                        </div>
                                        <div>{issue.message}</div>
                                    </div>
                                ))}
                            </div>
                        )}

                        {(spamReport?.triggerWords || spamReport?.spamWordsFound || []).length > 0 && (
                            <div className="mt-3 pt-3 border-t border-border">
                                <div className="text-[11px] font-semibold text-ink-muted uppercase tracking-wider mb-1.5">
                                    Spam Words Detected ({(spamReport?.triggerWords || spamReport?.spamWordsFound || []).length}):
                                </div>
                                <div className="flex flex-wrap gap-1">
                                    {(spamReport?.triggerWords || spamReport?.spamWordsFound || []).map((w, i) => (
                                        <span
                                            key={i}
                                            className="px-2 py-0.5 rounded bg-red-100 text-red-800 font-mono text-[10px] border border-red-200"
                                        >
                                            {w}
                                        </span>
                                    ))}
                                </div>
                            </div>
                        )}
                    </div>
                </aside>
            </div>

            {/* Action Buttons */}
            <div className="mt-8 flex justify-between items-center pt-4 border-t border-border">
                <button
                    type="button"
                    onClick={onBack}
                    className="btn btn-secondary text-xs px-4 py-2"
                >
                    Back
                </button>
                <button
                    type="button"
                    onClick={handleSubmit}
                    disabled={!subject.trim() || !body.trim()}
                    className="btn btn-primary text-xs px-6 py-2.5 disabled:opacity-40 disabled:cursor-not-allowed shadow-xs"
                >
                    Review Campaign
                </button>
            </div>

            {/* Saved Templates Modal */}
            <TemplateModal
                isOpen={isTemplateModalOpen}
                onClose={() => setIsTemplateModalOpen(false)}
                onSelectTemplate={(tpl) => {
                    setSubject(tpl.subject);
                    setBody(tpl.body);
                    if (editorRef.current) {
                        editorRef.current.innerHTML = tpl.body;
                    }
                    setIsTemplateModalOpen(false);
                }}
                currentSubject={subject}
                currentBody={body}
            />
        </div>
    );
};

export default EmailComposer;