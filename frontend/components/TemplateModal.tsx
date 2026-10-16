import React, { useState, useEffect } from 'react';
import { SavedTemplate } from '../types';
import apiClient from '../services/api';
import DOMPurify from 'dompurify';

interface TemplateModalProps {
    isOpen: boolean;
    onClose: () => void;
    onSelectTemplate: (template: { subject: string; body: string }) => void;
    currentSubject: string;
    currentBody: string;
}

const TemplateModal: React.FC<TemplateModalProps> = ({
    isOpen,
    onClose,
    onSelectTemplate,
    currentSubject,
    currentBody,
}) => {
    const [activeTab, setActiveTab] = useState<'browse' | 'save'>('browse');
    const [templates, setTemplates] = useState<SavedTemplate[]>([]);
    const [selectedCategory, setSelectedCategory] = useState<string>('All');
    const [selectedTemplate, setSelectedTemplate] = useState<SavedTemplate | null>(null);
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [saveName, setSaveName] = useState<string>('');
    const [saveCategory, setSaveCategory] = useState<string>('Cold Outreach');
    const [feedbackMsg, setFeedbackMsg] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

    // Fetch templates when modal opens
    useEffect(() => {
        if (isOpen) {
            loadTemplates();
            setFeedbackMsg(null);
        }
    }, [isOpen]);

    const loadTemplates = async () => {
        setIsLoading(true);
        try {
            const data = await apiClient.getTemplates();
            setTemplates(data.templates || []);
            if (data.templates && data.templates.length > 0) {
                setSelectedTemplate(data.templates[0]);
            }
        } catch (err: any) {
            console.error('Failed to load templates:', err);
        } finally {
            setIsLoading(false);
        }
    };

    const handleSaveTemplate = async (e: React.FormEvent) => {
        e.preventDefault();
        if (!saveName.trim()) {
            setFeedbackMsg({ type: 'error', text: 'Please enter a template name' });
            return;
        }
        if (!currentSubject.trim() || !currentBody.trim()) {
            setFeedbackMsg({ type: 'error', text: 'Current email has an empty subject or body' });
            return;
        }

        try {
            const res = await apiClient.createTemplate({
                name: saveName.trim(),
                category: saveCategory,
                subject: currentSubject,
                body: currentBody,
            });

            setTemplates(prev => [res.template, ...prev]);
            setSelectedTemplate(res.template);
            setSaveName('');
            setFeedbackMsg({ type: 'success', text: 'Template saved successfully' });
            setTimeout(() => {
                setActiveTab('browse');
                setFeedbackMsg(null);
            }, 1000);
        } catch (err: any) {
            setFeedbackMsg({ type: 'error', text: err.message || 'Failed to save template' });
        }
    };

    const handleDeleteTemplate = async (id: string, e: React.MouseEvent) => {
        e.stopPropagation();
        try {
            await apiClient.deleteTemplate(id);
            setTemplates(prev => prev.filter(t => t.id !== id));
            if (selectedTemplate?.id === id) {
                const remaining = templates.filter(t => t.id !== id);
                setSelectedTemplate(remaining.length > 0 ? remaining[0] : null);
            }
        } catch (err: any) {
            alert(err.message || 'Failed to delete template');
        }
    };

    if (!isOpen) return null;

    const categories = ['All', ...Array.from(new Set(templates.map(t => t.category)))];
    const filteredTemplates = selectedCategory === 'All'
        ? templates
        : templates.filter(t => t.category === selectedCategory);

    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-xs animate-fadeIn">
            <div className="surface-card w-full max-w-5xl max-h-[85vh] overflow-hidden flex flex-col border border-border shadow-xl bg-card rounded-none">
                {/* Header */}
                <div className="flex items-center justify-between px-5 py-3.5 border-b border-border bg-white">
                    <div className="flex items-center gap-2">
                        <div className="w-7 h-7 rounded-none bg-accent-light text-accent flex items-center justify-center">
                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" />
                            </svg>
                        </div>
                        <h2 className="text-base font-bold text-ink">Email Template Library</h2>
                    </div>

                    {/* Tabs */}
                    <div className="flex items-center gap-1 bg-neutral-100 p-0.5 rounded-none text-xs">
                        <button
                            type="button"
                            onClick={() => { setActiveTab('browse'); setFeedbackMsg(null); }}
                            className={`px-3 py-1 rounded-none font-medium transition-all ${
                                activeTab === 'browse'
                                    ? 'bg-white text-ink shadow-xs font-semibold'
                                    : 'text-ink-muted hover:text-ink'
                            }`}
                        >
                            Browse Templates
                        </button>
                        <button
                            type="button"
                            onClick={() => { setActiveTab('save'); setFeedbackMsg(null); }}
                            className={`px-3 py-1 rounded-none font-medium transition-all ${
                                activeTab === 'save'
                                    ? 'bg-white text-ink shadow-xs font-semibold'
                                    : 'text-ink-muted hover:text-ink'
                            }`}
                        >
                            Save Current
                        </button>
                    </div>

                    <button
                        onClick={onClose}
                        className="p-1.5 rounded-none text-ink-muted hover:text-ink hover:bg-neutral-100 transition-colors"
                    >
                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                        </svg>
                    </button>
                </div>

                {/* Content */}
                {activeTab === 'browse' ? (
                    <div className="flex-1 overflow-hidden flex flex-col md:flex-row min-h-[380px]">
                        {/* Sidebar: Categories & Template List */}
                        <div className="w-full md:w-5/12 border-r border-border flex flex-col bg-neutral-50/50">
                            {/* Category Filter Pills */}
                            <div className="p-3 border-b border-border overflow-x-auto flex gap-1.5 no-scrollbar">
                                {categories.map(cat => (
                                    <button
                                        key={cat}
                                        onClick={() => setSelectedCategory(cat)}
                                        className={`px-2.5 py-1 rounded-none text-[11px] font-medium whitespace-nowrap transition-colors ${
                                            selectedCategory === cat
                                                ? 'bg-ink text-white'
                                                : 'bg-white border border-border text-ink-secondary hover:text-ink'
                                        }`}
                                    >
                                        {cat}
                                    </button>
                                ))}
                            </div>

                            {/* Template List */}
                            <div className="flex-1 overflow-y-auto p-2 space-y-1.5">
                                {isLoading ? (
                                    <div className="p-4 text-center text-xs text-ink-muted">Loading templates...</div>
                                ) : filteredTemplates.length === 0 ? (
                                    <div className="p-4 text-center text-xs text-ink-muted">No templates found.</div>
                                ) : (
                                    filteredTemplates.map(t => {
                                        const isSelected = selectedTemplate?.id === t.id;
                                        return (
                                            <div
                                                key={t.id}
                                                onClick={() => setSelectedTemplate(t)}
                                                className={`p-2.5 rounded-none cursor-pointer border text-left transition-all relative group ${
                                                    isSelected
                                                        ? 'bg-white border-accent shadow-xs'
                                                        : 'bg-white/60 border-border/80 hover:bg-white hover:border-border'
                                                }`}
                                            >
                                                <div className="flex items-center justify-between gap-1 mb-1">
                                                    <span className="text-xs font-semibold text-ink truncate flex-1">{t.name}</span>
                                                    {t.isSystem ? (
                                                        <span className="px-1.5 py-0.5 rounded-none text-[10px] font-medium bg-neutral-100 text-ink-muted border border-neutral-200 flex-shrink-0">
                                                            Starter
                                                        </span>
                                                    ) : (
                                                        <button
                                                            type="button"
                                                            onClick={(e) => handleDeleteTemplate(t.id, e)}
                                                            className="text-ink-muted hover:text-red-600 p-0.5 opacity-0 group-hover:opacity-100 transition-opacity"
                                                            title="Delete Template"
                                                        >
                                                            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                                                            </svg>
                                                        </button>
                                                    )}
                                                </div>
                                                <p className="text-[11px] text-ink-muted truncate">{t.subject}</p>
                                            </div>
                                        );
                                    })
                                )}
                            </div>
                        </div>

                        {/* Preview Panel */}
                        <div className="w-full md:w-7/12 flex flex-col bg-white overflow-hidden">
                            {selectedTemplate ? (
                                <>
                                    <div className="p-4 border-b border-border bg-neutral-50/50">
                                        <div className="flex items-center justify-between gap-2 mb-1">
                                            <h3 className="text-sm font-bold text-ink truncate">{selectedTemplate.name}</h3>
                                            <span className="px-2 py-0.5 rounded-none text-[10px] font-medium bg-accent-light text-accent border border-neutral-200">
                                                {selectedTemplate.category}
                                            </span>
                                        </div>
                                        <div className="text-xs text-ink-secondary min-w-0">
                                            <span className="font-semibold text-ink-muted">Subject: </span>
                                            <span className="font-medium text-ink break-words">{selectedTemplate.subject}</span>
                                        </div>
                                    </div>

                                    {/* Body Preview */}
                                    <div className="flex-1 overflow-y-auto p-4 text-xs text-ink leading-relaxed break-words">
                                        <div
                                            dangerouslySetInnerHTML={{
                                                __html: DOMPurify.sanitize(selectedTemplate.body),
                                            }}
                                            className="prose prose-xs max-w-none break-words"
                                        />
                                    </div>

                                    {/* Action Bar */}
                                    <div className="p-3 border-t border-border bg-neutral-50/80 flex justify-end gap-2">
                                        <button
                                            type="button"
                                            onClick={() => {
                                                onSelectTemplate({
                                                    subject: selectedTemplate.subject,
                                                    body: selectedTemplate.body,
                                                });
                                                onClose();
                                            }}
                                            className="btn btn-primary text-xs px-4 py-2 font-semibold shadow-xs flex items-center gap-1.5"
                                        >
                                            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                                            </svg>
                                            Insert into Composer
                                        </button>
                                    </div>
                                </>
                            ) : (
                                <div className="flex-1 flex items-center justify-center p-6 text-xs text-ink-muted">
                                    Select a template on the left to preview
                                </div>
                            )}
                        </div>
                    </div>
                ) : (
                    /* Save Current Tab */
                    <form onSubmit={handleSaveTemplate} className="p-6 space-y-4 flex-1 overflow-y-auto">
                        {feedbackMsg && (
                            <div className={`p-3 rounded-none border text-xs font-medium ${
                                feedbackMsg.type === 'success'
                                    ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                                    : 'bg-red-50 border-red-200 text-red-800'
                            }`}>
                                {feedbackMsg.text}
                            </div>
                        )}

                        <div className="space-y-1.5">
                            <label className="block text-xs font-semibold text-ink">Template Name</label>
                            <input
                                type="text"
                                value={saveName}
                                onChange={(e) => setSaveName(e.target.value)}
                                placeholder="e.g. Senior Backend Referral Pitch"
                                className="input text-xs"
                                required
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="block text-xs font-semibold text-ink">Category</label>
                            <select
                                value={saveCategory}
                                onChange={(e) => setSaveCategory(e.target.value)}
                                className="input text-xs"
                            >
                                <option value="Cold Outreach">Cold Outreach</option>
                                <option value="Follow Up">Follow Up</option>
                                <option value="Referral">Referral</option>
                                <option value="Executive Pitch">Executive Pitch</option>
                                <option value="Custom">Custom</option>
                            </select>
                        </div>

                        <div className="p-3.5 rounded-none border border-border bg-neutral-50 text-xs space-y-2">
                            <div>
                                <span className="font-semibold text-ink-muted">Current Subject: </span>
                                <span className="font-medium text-ink break-words">{currentSubject || '(Empty Subject)'}</span>
                            </div>
                            <div>
                                <span className="font-semibold text-ink-muted">Content Preview: </span>
                                <div
                                    className="mt-1 p-2.5 rounded-none bg-white border border-border max-h-32 overflow-y-auto break-words"
                                    dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(currentBody || '<i>No body</i>') }}
                                />
                            </div>
                        </div>

                        <div className="flex justify-end pt-2">
                            <button
                                type="submit"
                                className="btn btn-primary text-xs px-5 py-2 font-semibold shadow-xs"
                            >
                                Save Template
                            </button>
                        </div>
                    </form>
                )}
            </div>
        </div>
    );
};

export default TemplateModal;
