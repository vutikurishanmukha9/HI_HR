import React, { useState, useRef, useEffect, useMemo, DragEvent } from 'react';
import DOMPurify from 'dompurify';
import { EmailTemplate, Recipient } from '../types';
import TemplateModal from './TemplateModal';
import { analyzeSpamScore } from '../utils/spamChecker';

interface EmailComposerProps {
    initialTemplate: EmailTemplate;
    onCompose: (template: EmailTemplate) => void;
    onBack: () => void;
    recipients?: Recipient[];
    availableTags?: string[];
}

const MAX_ATTACHMENT_SIZE = 10 * 1024 * 1024;
const MAX_SUBJECT_LENGTH = 200;

const EmailComposer: React.FC<EmailComposerProps> = ({
    initialTemplate,
    onCompose,
    onBack,
    availableTags = [],
    recipients = []
}) => {
    const [subject, setSubject] = useState(initialTemplate?.subject || '');
    const [body, setBody] = useState(initialTemplate?.body || '');
    const [attachments, setAttachments] = useState<File[]>(initialTemplate?.attachments || []);
    const editorRef = useRef<HTMLDivElement>(null);
    const subjectInputRef = useRef<HTMLInputElement>(null);
    const [showPlaceholder, setShowPlaceholder] = useState(false);
    const [isDraggingAttachment, setIsDraggingAttachment] = useState(false);
    const [attachmentError, setAttachmentError] = useState<string | null>(null);
    const [copiedTag, setCopiedTag] = useState<string | null>(null);
    const [insertedTag, setInsertedTag] = useState<string | null>(null);
    const [isTemplateModalOpen, setIsTemplateModalOpen] = useState(false);
    const [isVariableDropdownOpen, setIsVariableDropdownOpen] = useState(false);
    const [tagSearchQuery, setTagSearchQuery] = useState('');
    const [lastFocusedField, setLastFocusedField] = useState<'subject' | 'body'>('body');
    const [viewMode, setViewMode] = useState<'write' | 'preview'>('write');
    const [previewIndex, setPreviewIndex] = useState(0);

    const savedRangeRef = useRef<Range | null>(null);

    // Save selection range inside the editor for cursor insertion
    const saveSelection = () => {
        const selection = window.getSelection();
        if (selection && selection.rangeCount > 0 && editorRef.current) {
            const range = selection.getRangeAt(0);
            if (editorRef.current.contains(range.commonAncestorContainer)) {
                savedRangeRef.current = range.cloneRange();
            }
        }
    };

    // Resolve available tags dynamically from props, uploaded recipients, or defaults
    const resolvedTags = useMemo(() => {
        if (availableTags && availableTags.length > 0) {
            return availableTags;
        }
        if (recipients && recipients.length > 0) {
            const set = new Set<string>();
            recipients.forEach(r => {
                Object.keys(r).forEach(k => {
                    if (k !== 'email' && !k.startsWith('_') && k !== 'id') {
                        set.add(k);
                    }
                });
            });
            if (set.size > 0) return Array.from(set);
        }
        return ['name', 'company', 'role'];
    }, [availableTags, recipients]);

    // Active sample recipient for live preview and tag values
    const activeRecipient = useMemo(() => {
        if (recipients && recipients.length > 0) {
            return recipients[previewIndex] || recipients[0];
        }
        return {
            name: 'Sarah Connor',
            fullName: 'Sarah Connor',
            firstName: 'Sarah',
            company: 'Acme Technologies',
            companyName: 'Acme Technologies',
            role: 'Engineering Lead',
            jobTitle: 'Engineering Lead',
            location: 'San Francisco, CA',
            email: 'sarah.connor@example.com',
        };
    }, [recipients, previewIndex]);

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

    // Insert dynamic tag directly at the cursor position
    const insertTagAtCursor = (tag: string) => {
        const formattedTag = `{${tag}}`;

        if (lastFocusedField === 'subject' && subjectInputRef.current) {
            const input = subjectInputRef.current;
            const start = input.selectionStart ?? subject.length;
            const end = input.selectionEnd ?? subject.length;
            const nextSubject = subject.slice(0, start) + formattedTag + subject.slice(end);
            setSubject(nextSubject);
            setTimeout(() => {
                input.focus();
                input.setSelectionRange(start + formattedTag.length, start + formattedTag.length);
            }, 10);
            setInsertedTag(formattedTag);
            setTimeout(() => setInsertedTag(null), 1500);
            return;
        }

        // Insert into contentEditable body editor
        const editor = editorRef.current;
        if (!editor) return;

        editor.focus();
        const selection = window.getSelection();

        let range = savedRangeRef.current;
        if (!range || !editor.contains(range.commonAncestorContainer)) {
            if (selection && selection.rangeCount > 0) {
                const currentRange = selection.getRangeAt(0);
                if (editor.contains(currentRange.commonAncestorContainer)) {
                    range = currentRange;
                }
            }
        }

        if (!range) {
            range = document.createRange();
            range.selectNodeContents(editor);
            range.collapse(false);
        }

        if (selection) {
            selection.removeAllRanges();
            selection.addRange(range);
            range.deleteContents();
            const textNode = document.createTextNode(formattedTag + ' ');
            range.insertNode(textNode);
            range.setStartAfter(textNode);
            range.setEndAfter(textNode);
            selection.removeAllRanges();
            selection.addRange(range);
            savedRangeRef.current = range.cloneRange();
        }

        setBody(editor.innerHTML);
        setInsertedTag(formattedTag);
        setTimeout(() => setInsertedTag(null), 1500);
    };

    // Dynamic interpolator for live preview mode with smart alias resolution
    const interpolateContent = (content: string, highlight: boolean = true): string => {
        if (!content) return '';
        let result = content;
        const normalizedData: Record<string, string> = {};

        Object.entries(activeRecipient).forEach(([k, v]) => {
            if (v != null && typeof v !== 'object') {
                normalizedData[k] = String(v);
            }
        });

        // Smart alias bridges for cross-compatibility
        const nameVal = normalizedData.fullName || normalizedData.name || normalizedData.firstName || '';
        if (nameVal) {
            if (!normalizedData.name) normalizedData.name = nameVal;
            if (!normalizedData.fullName) normalizedData.fullName = nameVal;
            if (!normalizedData.firstName) normalizedData.firstName = normalizedData.firstName || nameVal.split(' ')[0];
        }

        const companyVal = normalizedData.companyName || normalizedData.company || normalizedData.organization || '';
        if (companyVal) {
            if (!normalizedData.company) normalizedData.company = companyVal;
            if (!normalizedData.companyName) normalizedData.companyName = companyVal;
        }

        const roleVal = normalizedData.jobTitle || normalizedData.role || normalizedData.title || normalizedData.position || '';
        if (roleVal) {
            if (!normalizedData.jobTitle) normalizedData.jobTitle = roleVal;
            if (!normalizedData.role) normalizedData.role = roleVal;
            if (!normalizedData.title) normalizedData.title = roleVal;
        }

        Object.entries(normalizedData).forEach(([k, v]) => {
            const regex = new RegExp(`\\{${k}\\}`, 'gi');
            const replacement = highlight
                ? `<span class="bg-amber-100 text-amber-900 px-1 py-0.5 rounded-none font-medium break-all">${v}</span>`
                : v;
            result = result.replace(regex, replacement);
        });

        return result;
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
        saveSelection();
    };

    const handleFormat = (command: string, value?: string) => {
        document.execCommand(command, false, value);
        editorRef.current?.focus();
        if (editorRef.current) setBody(editorRef.current.innerHTML);
        saveSelection();
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
        saveSelection();
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

    const ToolbarButton: React.FC<{ onClick: () => void; children: React.ReactNode; title: string }> = ({ onClick, children, title }) => (
        <button
            type="button"
            title={title}
            onClick={onClick}
            onMouseDown={(e) => {
                saveSelection();
                e.preventDefault();
            }}
            className="p-1.5 rounded-none text-ink-secondary hover:text-ink hover:bg-neutral-100 transition-colors"
        >
            {children}
        </button>
    );

    const scoreColors = useMemo(() => {
        if (spamReport.score >= 90) return { text: 'text-emerald-700', bg: 'bg-emerald-50', border: 'border-emerald-200', bar: 'bg-emerald-500' };
        if (spamReport.score >= 75) return { text: 'text-amber-700', bg: 'bg-amber-50', border: 'border-amber-200', bar: 'bg-amber-500' };
        return { text: 'text-red-700', bg: 'bg-red-50', border: 'border-red-200', bar: 'bg-red-500' };
    }, [spamReport.score]);

    // Filter tags based on user search input
    const filteredTags = useMemo(() => {
        if (!tagSearchQuery.trim()) return resolvedTags;
        const q = tagSearchQuery.toLowerCase();
        return resolvedTags.filter(t => t.toLowerCase().includes(q));
    }, [resolvedTags, tagSearchQuery]);

    return (
        <div className="max-w-6xl mx-auto">
            {/* Header */}
            <div className="text-center mb-6">
                <div className="inline-flex items-center justify-center w-12 h-12 rounded-none bg-accent-light border border-neutral-200 text-accent mb-3 shadow-xs">
                    <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                    </svg>
                </div>
                <h2 className="text-xl font-bold text-ink">Compose Your Email</h2>
                <p className="text-sm text-ink-muted mt-1">Craft personalized cold email copy using dynamic variables</p>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                {/* Main Editor Column */}
                <div className="lg:col-span-7 xl:col-span-8 space-y-5">
                    {/* Top Action Bar */}
                    <div className="flex flex-wrap justify-between items-center gap-2 p-3 bg-neutral-50 border border-border rounded-none">
                        <div className="flex items-center gap-2">
                            <button
                                type="button"
                                onClick={() => setIsTemplateModalOpen(true)}
                                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-none bg-card hover:bg-neutral-100 text-ink border border-neutral-200 shadow-xs transition-colors"
                            >
                                <svg className="w-4 h-4 text-accent" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" />
                                </svg>
                                Templates Library
                            </button>

                            {/* View Mode Toggle */}
                            <div className="inline-flex rounded-none border border-border bg-neutral-100 p-0.5 text-xs font-medium">
                                <button
                                    type="button"
                                    onClick={() => setViewMode('write')}
                                    className={`px-2.5 py-1 rounded-none transition-colors ${
                                        viewMode === 'write'
                                            ? 'bg-card text-ink font-semibold shadow-xs'
                                            : 'text-ink-muted hover:text-ink'
                                    }`}
                                >
                                    Write Template
                                </button>
                                <button
                                    type="button"
                                    onClick={() => setViewMode('preview')}
                                    className={`px-2.5 py-1 rounded-none transition-colors flex items-center gap-1 ${
                                        viewMode === 'preview'
                                            ? 'bg-card text-accent font-semibold shadow-xs'
                                            : 'text-ink-muted hover:text-ink'
                                    }`}
                                >
                                    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
                                    </svg>
                                    Live Preview
                                </button>
                            </div>
                        </div>

                        <div className="flex items-center gap-2">
                            <span className="text-[11px] text-ink-muted font-medium">Deliverability Score:</span>
                            <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-none text-xs font-bold border ${scoreColors.bg} ${scoreColors.text} ${scoreColors.border}`}>
                                <span className={`w-2 h-2 rounded-xs ${scoreColors.bar}`} />
                                {spamReport.score}/100 - {spamReport.rating}
                            </span>
                        </div>
                    </div>

                    {/* Live Preview Mode Card */}
                    {viewMode === 'preview' && (
                        <div className="surface-card border border-amber-200 bg-amber-50/40 rounded-none p-4 sm:p-5 space-y-3 animate-fadeIn">
                            <div className="flex items-center justify-between pb-3 border-b border-amber-200/60">
                                <div className="flex items-center gap-2">
                                    <span className="w-2 h-2 rounded-xs bg-emerald-500 animate-pulse" />
                                    <span className="text-xs font-bold text-amber-900 uppercase tracking-wider">
                                        Live Dynamic Simulation
                                    </span>
                                </div>
                                {recipients.length > 1 && (
                                    <div className="flex items-center gap-2 text-xs">
                                        <span className="text-ink-muted text-[11px]">
                                            Recipient {previewIndex + 1} of {recipients.length}
                                        </span>
                                        <div className="inline-flex rounded-none border border-border bg-card">
                                            <button
                                                type="button"
                                                onClick={() => setPreviewIndex(prev => Math.max(0, prev - 1))}
                                                disabled={previewIndex === 0}
                                                className="px-2 py-0.5 text-xs text-ink-secondary hover:text-ink disabled:opacity-30"
                                            >
                                                &larr;
                                            </button>
                                            <button
                                                type="button"
                                                onClick={() => setPreviewIndex(prev => Math.min(recipients.length - 1, prev + 1))}
                                                disabled={previewIndex >= recipients.length - 1}
                                                className="px-2 py-0.5 text-xs text-ink-secondary hover:text-ink disabled:opacity-30"
                                            >
                                                &rarr;
                                            </button>
                                        </div>
                                    </div>
                                )}
                            </div>

                            <div className="text-xs text-ink-muted min-w-0 truncate">
                                Simulating email for: <strong className="text-ink font-semibold">{activeRecipient.email || 'sample@recipient.com'}</strong>
                            </div>

                            <div className="bg-card border border-border rounded-none p-3 space-y-2">
                                <div className="text-xs font-semibold text-ink break-words">
                                    <span className="text-ink-muted mr-1">Subject:</span>
                                    <span dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(interpolateContent(subject)) || '<em class="text-ink-muted">Empty subject</em>' }} />
                                </div>
                                <div className="border-t border-border pt-2 text-xs leading-relaxed text-ink min-h-[140px] whitespace-pre-wrap break-words">
                                    <div dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(interpolateContent(body)) || '<em class="text-ink-muted">Empty body</em>' }} />
                                </div>
                            </div>

                            <p className="text-[11px] text-ink-muted">
                                Highlighted words represent dynamic variables populated from your spreadsheet.
                            </p>
                        </div>
                    )}

                    {/* Write Mode Form */}
                    <form onSubmit={handleSubmit} className={`space-y-4 ${viewMode === 'preview' ? 'hidden' : ''}`}>
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
                                ref={subjectInputRef}
                                type="text"
                                value={subject}
                                onChange={(e) => setSubject(e.target.value)}
                                onFocus={() => setLastFocusedField('subject')}
                                placeholder="Opportunity at {company}"
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
                            <div className="border border-border rounded-none overflow-hidden shadow-xs bg-card focus-within:border-accent focus-within:ring-2 focus-within:ring-neutral-200 transition-all">
                                {/* Toolbar */}
                                <div className="flex flex-wrap items-center gap-1 px-3 py-2 border-b border-border bg-neutral-50">
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
                                    <div className="w-px h-4 mx-1 bg-border" />

                                    {/* Insert Variable Dropdown */}
                                    <div className="relative">
                                        <button
                                            type="button"
                                            title="Insert dynamic personalization tag at cursor"
                                            onClick={() => setIsVariableDropdownOpen(!isVariableDropdownOpen)}
                                            onMouseDown={(e) => {
                                                saveSelection();
                                                e.preventDefault();
                                            }}
                                            className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-none text-xs font-medium text-ink bg-neutral-100 hover:bg-neutral-200 border border-neutral-300 transition-colors"
                                        >
                                            <span className="font-mono font-bold text-accent">{`{ }`}</span>
                                            <span>Insert Variable</span>
                                            <svg className="w-3 h-3 text-ink-muted" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                                            </svg>
                                        </button>

                                        {isVariableDropdownOpen && (
                                            <div className="absolute left-0 top-full mt-1 w-64 bg-card border border-border rounded-none shadow-lg z-30 py-1.5 animate-fadeIn">
                                                <div className="px-3 py-1.5 text-[10px] uppercase font-bold text-ink-muted tracking-wider border-b border-border flex justify-between items-center">
                                                    <span>Available Columns ({resolvedTags.length})</span>
                                                    <span className="text-emerald-700 font-sans font-normal lowercase">dynamic</span>
                                                </div>
                                                <div className="max-h-52 overflow-y-auto py-1">
                                                    {resolvedTags.map(tag => {
                                                        const sampleVal = (activeRecipient as any)[tag] || '';
                                                        return (
                                                            <button
                                                                key={tag}
                                                                type="button"
                                                                onClick={() => {
                                                                    insertTagAtCursor(tag);
                                                                    setIsVariableDropdownOpen(false);
                                                                }}
                                                                className="w-full text-left px-3 py-1.5 hover:bg-neutral-50 flex items-center justify-between text-xs transition-colors"
                                                            >
                                                                <span className="font-mono font-semibold text-accent">{`{${tag}}`}</span>
                                                                {sampleVal && (
                                                                    <span className="text-[11px] text-ink-muted truncate max-w-[110px]" title={String(sampleVal)}>
                                                                        {String(sampleVal)}
                                                                    </span>
                                                                )}
                                                            </button>
                                                        );
                                                    })}
                                                </div>
                                            </div>
                                        )}
                                    </div>
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
                                        onFocus={() => {
                                            setLastFocusedField('body');
                                            saveSelection();
                                        }}
                                        onKeyUp={saveSelection}
                                        onMouseUp={saveSelection}
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
                            <div className="p-3 rounded-none bg-danger-light border border-red-200 text-danger text-xs font-medium animate-fadeIn break-words">
                                {attachmentError}
                            </div>
                        )}
                        <div
                            className={`p-5 rounded-none text-center cursor-pointer border-2 border-dashed transition-all ${
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
                                        className="flex items-center justify-between p-2.5 rounded-none bg-neutral-50 border border-border text-xs"
                                    >
                                        <div className="flex items-center gap-2 min-w-0 flex-1 pr-2">
                                            <svg className="w-4 h-4 text-accent flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                                                <path fillRule="evenodd" d="M8 4a3 3 0 00-3 3v4a3 3 0 006 0V7a1 1 0 112 0v4a5 5 0 01-10 0V7a5 5 0 0110 0v4a1 1 0 11-2 0V7a3 3 0 00-3-3z" clipRule="evenodd" />
                                            </svg>
                                            <span className="font-medium text-ink truncate max-w-[240px]" title={file.name}>{file.name}</span>
                                            <span className="text-ink-muted text-[10px] flex-shrink-0">({(file.size / 1024).toFixed(0)} KB)</span>
                                        </div>
                                        <button
                                            type="button"
                                            onClick={() => removeAttachment(file.name)}
                                            className="text-danger hover:text-red-700 font-medium text-[11px] flex-shrink-0"
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
                <aside className="lg:col-span-5 xl:col-span-4 space-y-5">
                    {/* Personalization Tags */}
                    <div className="surface-card p-4 sm:p-5 border border-border shadow-xs rounded-none">
                        <div className="flex items-center justify-between mb-1.5">
                            <div className="flex items-center gap-2">
                                <div className="w-7 h-7 rounded-none bg-neutral-100 border border-neutral-200 flex items-center justify-center text-accent">
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 7h.01M7 3h5c.512 0 1.024.195 1.414.586l7 7a2 2 0 010 2.828l-7 7a2 2 0 01-2.828 0l-7-7A1.994 1.994 0 013 12V7a4 4 0 014-4z" />
                                    </svg>
                                </div>
                                <h3 className="font-semibold text-sm text-ink">Personalization Tags</h3>
                            </div>
                            <span className="text-[10px] font-semibold px-2 py-0.5 rounded-none bg-emerald-50 text-emerald-800 border border-emerald-200">
                                {recipients.length > 0 ? `${resolvedTags.length} Detected` : 'Dynamic'}
                            </span>
                        </div>
                        <p className="text-xs text-ink-muted mb-3">Click any variable to insert directly into your email at cursor</p>

                        {/* Search filter if more than 4 tags */}
                        {resolvedTags.length > 4 && (
                            <div className="mb-2">
                                <input
                                    type="text"
                                    value={tagSearchQuery}
                                    onChange={(e) => setTagSearchQuery(e.target.value)}
                                    placeholder="Search columns..."
                                    className="input py-1 text-xs w-full rounded-none"
                                />
                            </div>
                        )}

                        {filteredTags.length === 0 ? (
                            <div className="text-xs p-3.5 rounded-none text-center bg-neutral-50 border border-border text-ink-muted">
                                No matching tags found.
                            </div>
                        ) : (
                            <div className="space-y-2 max-h-[250px] overflow-y-auto pr-1">
                                {filteredTags.map(tag => {
                                    const formattedTag = `{${tag}}`;
                                    const isCopied = copiedTag === formattedTag;
                                    const isInserted = insertedTag === formattedTag;
                                    const sampleVal = (activeRecipient as any)[tag];

                                    return (
                                        <div
                                            key={tag}
                                            className="p-2.5 rounded-none bg-neutral-50 border border-border hover:border-neutral-300 transition-all flex flex-col gap-1.5"
                                        >
                                            <div className="flex items-center justify-between gap-2 min-w-0">
                                                <span className="font-mono font-semibold text-xs text-ink truncate min-w-0" title={formattedTag}>
                                                    {formattedTag}
                                                </span>
                                                <div className="flex items-center gap-1 flex-shrink-0">
                                                    {/* Quick Insert Button */}
                                                    <button
                                                        type="button"
                                                        onClick={() => insertTagAtCursor(tag)}
                                                        title="Insert into email body or subject"
                                                        className={`px-2 py-0.5 rounded-none text-[11px] font-medium transition-all ${
                                                            isInserted
                                                                ? 'bg-emerald-600 text-white'
                                                                : 'bg-accent text-white hover:bg-neutral-800'
                                                        }`}
                                                    >
                                                        {isInserted ? 'Inserted' : '+ Insert'}
                                                    </button>
                                                    {/* Quick Copy Button */}
                                                    <button
                                                        type="button"
                                                        onClick={() => copyToClipboard(formattedTag)}
                                                        title="Copy tag to clipboard"
                                                        className={`p-1 rounded-none text-[11px] border transition-all ${
                                                            isCopied
                                                                ? 'bg-emerald-50 border-emerald-300 text-emerald-700'
                                                                : 'bg-card border-border hover:bg-neutral-100 text-ink-secondary'
                                                        }`}
                                                    >
                                                        {isCopied ? (
                                                            <svg className="w-3.5 h-3.5 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M5 13l4 4L19 7" />
                                                            </svg>
                                                        ) : (
                                                            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" />
                                                            </svg>
                                                        )}
                                                    </button>
                                                </div>
                                            </div>
                                            {/* Sample Value Preview */}
                                            {sampleVal && (
                                                <div className="text-[11px] text-ink-muted truncate max-w-full">
                                                    Sample: <span className="text-ink font-medium">"{String(sampleVal)}"</span>
                                                </div>
                                            )}
                                        </div>
                                    );
                                })}
                            </div>
                        )}
                        <p className="text-[11px] text-ink-muted mt-2 pt-2 border-t border-border">
                            Every column in your uploaded spreadsheet is dynamically mapped and replaced per recipient at send time.
                        </p>
                    </div>

                    {/* Deliverability & Spam Insights */}
                    <div className="surface-card p-4 sm:p-5 border border-border shadow-xs rounded-none">
                        <div className="flex items-center justify-between mb-2">
                            <div className="flex items-center gap-2">
                                <div className="w-7 h-7 rounded-none bg-neutral-100 border border-neutral-200 flex items-center justify-center text-accent">
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
                                    </svg>
                                </div>
                                <h3 className="font-semibold text-sm text-ink">Deliverability & Spam Engine</h3>
                            </div>
                            <span className={`text-xs font-bold px-2 py-0.5 rounded-none border ${scoreColors.bg} ${scoreColors.text} ${scoreColors.border}`}>
                                {spamReport.score}/100 {spamReport.rating}
                            </span>
                        </div>

                        {/* Meter Bar */}
                        <div className="w-full bg-neutral-100 rounded-none h-2 my-3 overflow-hidden border border-neutral-200">
                            <div
                                className={`h-full transition-all duration-300 ${scoreColors.bar}`}
                                style={{ width: `${spamReport.score}%` }}
                            />
                        </div>

                        {/* Rectangular Quick Deliverability Metrics */}
                        {spamReport.metrics && (
                            <div className="grid grid-cols-2 gap-2 my-3">
                                <div className="p-2 bg-neutral-50 border border-neutral-200 rounded-none">
                                    <div className="text-[10px] uppercase font-semibold text-ink-muted tracking-wider">Word Count</div>
                                    <div className="text-xs font-bold text-ink mt-0.5">{spamReport.metrics.wordCount} words</div>
                                </div>
                                <div className="p-2 bg-neutral-50 border border-neutral-200 rounded-none">
                                    <div className="text-[10px] uppercase font-semibold text-ink-muted tracking-wider">Read Time</div>
                                    <div className="text-xs font-bold text-ink mt-0.5">~{spamReport.metrics.readingTimeSeconds}s</div>
                                </div>
                                <div className="p-2 bg-neutral-50 border border-neutral-200 rounded-none">
                                    <div className="text-[10px] uppercase font-semibold text-ink-muted tracking-wider">Links Count</div>
                                    <div className="text-xs font-bold text-ink mt-0.5">{spamReport.metrics.linkCount}</div>
                                </div>
                                <div className="p-2 bg-neutral-50 border border-neutral-200 rounded-none">
                                    <div className="text-[10px] uppercase font-semibold text-ink-muted tracking-wider">Personalized</div>
                                    <div className={`text-xs font-bold mt-0.5 ${spamReport.metrics.hasPersonalization ? 'text-emerald-700' : 'text-amber-700'}`}>
                                        {spamReport.metrics.hasPersonalization ? 'Active' : 'Missing Tags'}
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* Issues list */}
                        {spamReport.issues.length === 0 ? (
                            <div className="p-3 rounded-none bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs flex items-center gap-2">
                                <svg className="w-4 h-4 text-emerald-600 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                                </svg>
                                <span>High inbox placement potential. No spam triggers detected.</span>
                            </div>
                        ) : (
                            <div className="space-y-2.5 max-h-[280px] overflow-y-auto pr-2">
                                {spamReport.issues.map((issue, idx) => (
                                    <div
                                        key={idx}
                                        className={`p-3 rounded-none border text-xs leading-relaxed ${
                                            issue.severity === 'high'
                                                ? 'bg-red-50/70 border-red-200 text-red-950'
                                                : issue.severity === 'medium'
                                                ? 'bg-amber-50/70 border-amber-200 text-amber-950'
                                                : 'bg-neutral-50 border-neutral-200 text-ink-secondary'
                                        }`}
                                    >
                                        <div className="flex items-start justify-between gap-2 mb-1.5">
                                            <span className="font-bold text-[11px] uppercase tracking-wider text-ink break-words">
                                                {issue.type}
                                            </span>
                                            <span className="text-[9px] uppercase font-bold tracking-wider px-1.5 py-0.5 bg-white border border-current rounded-none flex-shrink-0 whitespace-nowrap">
                                                {issue.severity === 'high' ? 'High Risk' : issue.severity === 'medium' ? 'Notice' : 'Suggestion'}
                                            </span>
                                        </div>
                                        <div className="text-[11px] leading-relaxed break-words">{issue.message}</div>
                                        {issue.suggestion && (
                                            <div className="mt-2 pt-1.5 border-t border-current/20 text-[10.5px] leading-relaxed break-words">
                                                <span className="font-semibold">Tip:</span> {issue.suggestion}
                                            </div>
                                        )}
                                    </div>
                                ))}
                            </div>
                        )}

                        {/* Detected Spam Trigger Words */}
                        {(spamReport?.detectedWords && spamReport.detectedWords.length > 0) ? (
                            <div className="mt-3 pt-3 border-t border-border">
                                <div className="text-[11px] font-bold text-ink-muted uppercase tracking-wider mb-2">
                                    Spam Words Detected ({spamReport.detectedWords.length}):
                                </div>
                                <div className="space-y-2">
                                    {spamReport.detectedWords.map((item, idx) => (
                                        <div
                                            key={idx}
                                            className={`p-2.5 border rounded-none text-xs ${
                                                item.severity === 'high'
                                                    ? 'bg-red-50 border-red-200 text-red-950'
                                                    : 'bg-amber-50 border-amber-200 text-amber-950'
                                            }`}
                                        >
                                            <div className="flex items-center justify-between gap-2 mb-1.5">
                                                <span className="font-mono font-bold px-1.5 py-0.5 bg-white border border-current rounded-none text-xs break-all">
                                                    {item.word}
                                                </span>
                                                <div className="flex items-center gap-1 flex-shrink-0">
                                                    <span className="text-[9px] uppercase tracking-wider px-1 py-0.5 bg-white border border-current rounded-none whitespace-nowrap">
                                                        {item.location}
                                                    </span>
                                                    <span className="text-[9px] uppercase tracking-wider px-1.5 py-0.5 bg-white border border-current rounded-none font-bold whitespace-nowrap">
                                                        {item.severity === 'high' ? 'High Risk' : 'Caution'}
                                                    </span>
                                                </div>
                                            </div>
                                            <div className="text-[11px] leading-relaxed text-ink-secondary mt-1 break-words">
                                                <span className="font-semibold text-ink">Fix:</span> {item.suggestion}
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            </div>
                        ) : (spamReport?.triggerWords || spamReport?.spamWordsFound || []).length > 0 ? (
                            <div className="mt-3 pt-3 border-t border-border">
                                <div className="text-[11px] font-semibold text-ink-muted uppercase tracking-wider mb-1.5">
                                    Spam Words Detected ({(spamReport?.triggerWords || spamReport?.spamWordsFound || []).length}):
                                </div>
                                <div className="flex flex-wrap gap-1">
                                    {(spamReport?.triggerWords || spamReport?.spamWordsFound || []).map((w, i) => (
                                        <span
                                            key={i}
                                            className="px-2 py-0.5 rounded-none bg-red-100 text-red-800 font-mono text-[10px] border border-red-200"
                                        >
                                            {w}
                                        </span>
                                    ))}
                                </div>
                            </div>
                        ) : null}
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