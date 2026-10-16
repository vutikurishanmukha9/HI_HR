import React, { useState, useCallback, DragEvent } from 'react';
import * as XLSX from 'xlsx';
import { Recipient, DomainMxResult } from '../types';
import { apiClient } from '../services/api';

interface RecipientUploaderProps {
  onUpload: (recipients: Recipient[]) => void;
  onBack: () => void;
}

const findValue = (row: any, potentialKeys: string[]): string => {
  const rowKeys = Object.keys(row);
  for (const pKey of potentialKeys) {
    const foundKey = rowKeys.find(rKey => rKey.trim().toLowerCase() === pKey.trim().toLowerCase());
    if (foundKey && row[foundKey] != null) {
      return String(row[foundKey]).trim();
    }
  }
  return '';
};

// Helper to sanitize header strings to camelCase (e.g. "First Name" -> "firstName")
const sanitizeKey = (key: string): string => {
  const cleaned = key.replace(/[^a-zA-Z0-9]/g, ' ').trim();
  return cleaned.split(/\s+/).map((word, index) => {
    if (index === 0) return word.toLowerCase();
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join('');
};

const RecipientUploader: React.FC<RecipientUploaderProps> = ({ onUpload, onBack }) => {
  const [recipients, setRecipients] = useState<Recipient[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [warning, setWarning] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [isValidatingMx, setIsValidatingMx] = useState(false);
  const [mxResults, setMxResults] = useState<Record<string, DomainMxResult> | null>(null);
  const [mxSummary, setMxSummary] = useState<{
    totalChecked: number;
    validCount: number;
    invalidCount: number;
    disposableCount: number;
    roleCount: number;
  } | null>(null);

  const MAX_FILE_SIZE = 5 * 1024 * 1024;

  const processFile = useCallback((file: File) => {
    if (file.size > MAX_FILE_SIZE) {
      setError(`File size exceeds 5MB limit. Please use a smaller file.`);
      setLoading(false);
      return;
    }

    setLoading(true);
    setFileName(file.name);
    setError(null);
    setWarning(null);
    setRecipients([]);
    setMxResults(null);
    setMxSummary(null);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);

        const parsedRecipients: Recipient[] = json.map((row: any): Recipient => {
          // Find the email using common patterns before sanitization
          const email = findValue(row, ['Email address', 'Email', 'E-mail', 'Contact Email']);
          
          const sanitizedRow: Record<string, any> = {};
          Object.keys(row).forEach(key => {
            const camelKey = sanitizeKey(key);
            if (camelKey) {
              sanitizedRow[camelKey] = String(row[key]).trim();
            }
          });

          return {
            ...sanitizedRow,
            email: email || sanitizedRow.email || '' // explicitly map the email property
          };
        }).filter((r: Recipient) => r.email); // Only require email to be valid

        if (parsedRecipients.length === 0) {
          setError('No valid recipients found. Ensure your file has Name, Email, and Company columns.');
        } else {
          const emailSet = new Set<string>();
          const duplicates: string[] = [];
          parsedRecipients.forEach(r => {
            const email = r.email.toLowerCase();
            if (emailSet.has(email)) {
              duplicates.push(email);
            } else {
              emailSet.add(email);
            }
          });

          if (duplicates.length > 0) {
            setWarning(`Found ${duplicates.length} duplicate email(s). Only the first occurrence will be used.`);
          }

          const uniqueRecipients = parsedRecipients.filter((r, index, self) =>
            index === self.findIndex(t => t.email.toLowerCase() === r.email.toLowerCase())
          );
          setRecipients(uniqueRecipients);
        }
      } catch (err) {
        setError('Failed to parse the Excel file. Please ensure it is a valid .xlsx or .xls file.');
        console.error(err);
      } finally {
        setLoading(false);
      }
    };
    reader.onerror = () => {
      setError('Failed to read the file.');
      setLoading(false);
    };
    reader.readAsBinaryString(file);
  }, [MAX_FILE_SIZE]);

  const handleFileChange = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) processFile(file);
  }, [processFile]);

  const handleDragEnter = (e: DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); };
  const handleDragLeave = (e: DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); };
  const handleDragOver = (e: DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); };
  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.csv'))) {
      processFile(file);
    } else {
      setError('Please drop a valid file (.xlsx, .xls, .csv).');
    }
  };

  const handleValidateDomains = async () => {
    if (recipients.length === 0) return;
    setIsValidatingMx(true);
    setError(null);
    try {
      const emails = recipients.map(r => r.email);
      const res = await apiClient.validateDomains(emails);
      const resultMap: Record<string, DomainMxResult> = {};
      if (res && res.results) {
        Object.entries(res.results).forEach(([key, item]) => {
          if (item) {
            resultMap[key.toLowerCase()] = item;
            if (item.domain) {
              resultMap[item.domain.toLowerCase()] = item;
            }
            if (item.email) {
              resultMap[item.email.toLowerCase()] = item;
            }
          }
        });
      }
      let validCount = 0;
      let invalidCount = 0;
      let disposableCount = 0;
      let roleCount = 0;

      recipients.forEach(r => {
        const emailKey = (r.email || '').trim().toLowerCase();
        const domainKey = (r.email.split('@')[1] || '').trim().toLowerCase();
        const dRes = resultMap[emailKey] || resultMap[domainKey];
        if (dRes && dRes.hasMx) {
          validCount++;
        } else {
          invalidCount++;
        }
        if (dRes && dRes.isDisposable) {
          disposableCount++;
        }
        if (dRes && dRes.isRoleAccount) {
          roleCount++;
        }
      });

      setMxResults(resultMap);
      setMxSummary({
        totalChecked: recipients.length,
        validCount,
        invalidCount,
        disposableCount,
        roleCount
      });
    } catch (err: any) {
      setError(err.message || 'Failed to verify recipient domain DNS and deliverability records');
    } finally {
      setIsValidatingMx(false);
    }
  };

  const handleRemoveInvalidRecipients = () => {
    if (!mxResults) return;
    const filtered = recipients.filter(r => {
      const emailKey = (r.email || '').trim().toLowerCase();
      const domainKey = (r.email.split('@')[1] || '').trim().toLowerCase();
      const res = mxResults[emailKey] || mxResults[domainKey];
      return res ? res.hasMx : true;
    });
    setRecipients(filtered);
    setMxSummary(prev => prev ? { ...prev, invalidCount: 0 } : null);
  };

  const handleRemoveDisposableRecipients = () => {
    if (!mxResults) return;
    const filtered = recipients.filter(r => {
      const emailKey = (r.email || '').trim().toLowerCase();
      const domainKey = (r.email.split('@')[1] || '').trim().toLowerCase();
      const res = mxResults[emailKey] || mxResults[domainKey];
      return res ? !res.isDisposable : true;
    });
    setRecipients(filtered);
    setMxSummary(prev => prev ? { ...prev, disposableCount: 0 } : null);
  };

  const handleCleanAll = () => {
    if (!mxResults) return;
    const filtered = recipients.filter(r => {
      const emailKey = (r.email || '').trim().toLowerCase();
      const domainKey = (r.email.split('@')[1] || '').trim().toLowerCase();
      const res = mxResults[emailKey] || mxResults[domainKey];
      if (!res) return true;
      return res.hasMx && !res.isDisposable;
    });
    setRecipients(filtered);
    setMxSummary(prev => prev ? { ...prev, invalidCount: 0, disposableCount: 0 } : null);
  };

  const handleSubmit = () => {
    if (recipients.length > 0) onUpload(recipients);
  };

  return (
    <div className="max-w-6xl mx-auto">
      {/* Header */}
      <div className="text-center mb-6">
        <div className="inline-flex items-center justify-center w-12 h-12 rounded-none bg-accent-light border border-neutral-200 text-accent mb-3 shadow-xs">
          <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
          </svg>
        </div>
        <h2 className="text-xl font-bold text-ink">Upload Outreach List</h2>
        <p className="text-sm text-ink-muted mt-1">Upload a CSV or Excel spreadsheet containing your prospective contacts</p>
      </div>

      {/* Drop Zone */}
      <div
        className={`relative p-8 rounded-none text-center cursor-pointer border-2 border-dashed transition-all duration-150 ${
          isDragging
            ? 'border-accent bg-neutral-100 scale-[1.01]'
            : 'border-neutral-300 hover:border-accent/80 bg-neutral-50 hover:bg-neutral-100/60'
        }`}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDragOver={handleDragOver}
        onDrop={handleDrop}
        onClick={() => document.getElementById('file-upload')?.click()}
      >
        <div className="flex flex-col items-center gap-3">
          {loading ? (
            <div className="w-10 h-10 border-3 border-accent border-t-transparent rounded-full animate-spin" />
          ) : (
            <div className="w-14 h-14 rounded-none bg-white border border-border shadow-xs flex items-center justify-center text-accent">
              <svg className="w-7 h-7" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.8} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
              </svg>
            </div>
          )}

          <div className="max-w-md mx-auto w-full px-2">
            <p className="font-semibold text-ink text-sm truncate" title={fileName}>
              {fileName ? fileName : 'Choose a file or drag & drop here'}
            </p>
            <p className="text-xs text-ink-muted mt-1">Supports CSV, XLSX, XLS up to 5MB</p>
          </div>

          <input id="file-upload" type="file" className="sr-only" onChange={handleFileChange} accept=".xlsx,.xls,.csv" />
        </div>
      </div>

      {/* Messages */}
      {error && (
        <div className="mt-4 p-3.5 rounded-none bg-danger-light border border-red-200 text-danger animate-fadeIn">
          <p className="text-xs font-medium flex items-center gap-2 break-words">
            <svg className="w-4 h-4 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
            </svg>
            <span className="min-w-0 break-words">{error}</span>
          </p>
        </div>
      )}

      {warning && (
        <div className="mt-4 p-3.5 rounded-none bg-amber-50 border border-amber-200 text-amber-900 animate-fadeIn">
          <p className="text-xs font-medium flex items-center gap-2 break-words">
            <svg className="w-4 h-4 text-amber-600 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
            </svg>
            <span className="min-w-0 break-words">{warning}</span>
          </p>
        </div>
      )}

      {/* Recipients Table */}
      {recipients.length > 0 && (
        <div className="mt-6 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-3 p-3 rounded-none bg-neutral-50 border border-border">
            <div className="flex items-center gap-2">
              <span className="badge badge-success font-semibold rounded-none">{recipients.length} recipients</span>
              <span className="text-xs text-ink-muted">ready for personalization</span>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {mxSummary && (
                <div className="flex flex-wrap items-center gap-1.5 text-xs mr-1">
                  <span className="px-2 py-0.5 rounded-none bg-emerald-50 text-emerald-700 border border-emerald-200 font-medium">
                    {mxSummary.validCount} MX OK
                  </span>
                  {mxSummary.invalidCount > 0 && (
                    <span className="px-2 py-0.5 rounded-none bg-red-50 text-red-700 border border-red-200 font-medium">
                      {mxSummary.invalidCount} Dead Domains
                    </span>
                  )}
                  {mxSummary.disposableCount > 0 && (
                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-none bg-amber-50 text-amber-700 border border-amber-200 font-medium">
                      <svg className="w-3 h-3 text-amber-500 flex-shrink-0" viewBox="0 0 20 20" fill="currentColor">
                        <path fillRule="evenodd" d="M8.485 2.495c.673-1.167 2.357-1.167 3.03 0l6.28 10.875c.673 1.167-.17 2.625-1.516 2.625H3.72c-1.347 0-2.189-1.458-1.515-2.625L8.485 2.495zM10 5a.75.75 0 01.75.75v3.5a.75.75 0 01-1.5 0v-3.5A.75.75 0 0110 5zm0 9a1 1 0 100-2 1 1 0 000 2z" clipRule="evenodd" />
                      </svg>
                      {mxSummary.disposableCount} Temp Mail
                    </span>
                  )}
                  {mxSummary.roleCount > 0 && (
                    <span className="px-2 py-0.5 rounded-none bg-purple-50 text-purple-700 border border-purple-200 font-medium" title="Generic role inboxes (info@, sales@, support@, etc.)">
                      {mxSummary.roleCount} Role Inboxes
                    </span>
                  )}
                </div>
              )}

              {mxSummary && (mxSummary.invalidCount > 0 || mxSummary.disposableCount > 0) && (
                <div className="flex items-center gap-1.5">
                  {mxSummary.disposableCount > 0 && (
                    <button
                      type="button"
                      onClick={handleRemoveDisposableRecipients}
                      className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-none text-xs font-semibold bg-amber-50 text-amber-700 hover:bg-amber-100 border border-amber-200 transition-colors"
                      title="Remove temporary and disposable burner email addresses"
                    >
                      <svg className="w-3.5 h-3.5 text-amber-500 flex-shrink-0" viewBox="0 0 20 20" fill="currentColor">
                        <path fillRule="evenodd" d="M8.485 2.495c.673-1.167 2.357-1.167 3.03 0l6.28 10.875c.673 1.167-.17 2.625-1.516 2.625H3.72c-1.347 0-2.189-1.458-1.515-2.625L8.485 2.495zM10 5a.75.75 0 01.75.75v3.5a.75.75 0 01-1.5 0v-3.5A.75.75 0 0110 5zm0 9a1 1 0 100-2 1 1 0 000 2z" clipRule="evenodd" />
                      </svg>
                      Exclude {mxSummary.disposableCount} Temp
                    </button>
                  )}
                  {mxSummary.invalidCount > 0 && (
                    <button
                      type="button"
                      onClick={handleRemoveInvalidRecipients}
                      className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-none text-xs font-semibold bg-red-50 text-red-700 hover:bg-red-100 border border-red-200 transition-colors"
                      title="Remove dead domains with no active mail server"
                    >
                      Exclude {mxSummary.invalidCount} Dead
                    </button>
                  )}
                  {mxSummary.invalidCount > 0 && mxSummary.disposableCount > 0 && (
                    <button
                      type="button"
                      onClick={handleCleanAll}
                      className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-none text-xs font-semibold bg-card text-ink hover:bg-neutral-100 border border-border shadow-xs transition-colors"
                      title="Remove both dead domains and disposable emails at once"
                    >
                      Clean List
                    </button>
                  )}
                </div>
              )}

              <button
                type="button"
                onClick={handleValidateDomains}
                disabled={isValidatingMx}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-none text-xs font-semibold bg-card hover:bg-neutral-100 text-ink border border-neutral-200 shadow-xs transition-colors disabled:opacity-50"
              >
                {isValidatingMx ? (
                  <>
                    <div className="w-3.5 h-3.5 border-2 border-accent border-t-transparent rounded-full animate-spin" />
                    <span>Verifying MX & Temp...</span>
                  </>
                ) : (
                  <>
                    <svg className="w-3.5 h-3.5 text-accent" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
                    </svg>
                    <span>Verify Domains (DNS MX)</span>
                  </>
                )}
              </button>
            </div>
          </div>

          <div className="border border-border rounded-none overflow-hidden shadow-xs bg-card">
            <div className="max-h-64 overflow-auto">
              <table className="table w-full">
                <thead>
                  <tr>
                    {Object.keys(recipients[0] || {}).slice(0, 5).map((key) => (
                      <th key={key} className="text-[11px] whitespace-nowrap">
                        {key}
                      </th>
                    ))}
                    <th className="text-[11px] whitespace-nowrap">Domain MX Status</th>
                    <th className="text-[11px] whitespace-nowrap">Mail Type</th>
                  </tr>
                </thead>
                <tbody>
                  {recipients.slice(0, 10).map((recipient, index) => {
                    const emailKey = (recipient.email || '').trim().toLowerCase();
                    const domain = (recipient.email ? recipient.email.split('@')[1] : '')?.toLowerCase();
                    const mxInfo = mxResults ? (mxResults[emailKey] || mxResults[domain]) : null;

                    return (
                      <tr key={index} className="hover:bg-neutral-50 transition-colors">
                        {Object.keys(recipients[0] || {}).slice(0, 5).map((key) => {
                          const val = recipient[key] || '—';
                          return (
                            <td 
                              key={key} 
                              className={`text-xs truncate max-w-[200px] break-all ${key === 'email' ? 'font-mono text-ink-secondary' : 'text-ink'}`}
                              title={typeof val === 'string' ? val : String(val)}
                            >
                              {val}
                            </td>
                          );
                        })}
                        <td className="text-xs whitespace-nowrap">
                          {mxInfo ? (
                            mxInfo.hasMx ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-none text-[10px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                                <svg className="w-3 h-3 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M5 13l4 4L19 7" />
                                </svg>
                                MX Active
                              </span>
                            ) : (
                              <span
                                className="inline-flex items-center gap-1 px-2 py-0.5 rounded-none text-[10px] font-semibold bg-red-50 text-red-700 border border-red-200"
                                title={mxInfo.error || 'No valid MX records'}
                              >
                                <svg className="w-3 h-3 text-red-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" />
                                </svg>
                                Dead Domain
                              </span>
                            )
                          ) : (
                            <span className="text-[11px] text-ink-muted">Not verified</span>
                          )}
                        </td>
                        <td className="text-xs whitespace-nowrap">
                          {mxInfo ? (
                            mxInfo.isDisposable ? (
                              <span
                                className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-none text-[10px] font-semibold bg-amber-50 text-amber-700 border border-amber-200 whitespace-nowrap shadow-2xs"
                                title={mxInfo.disposableReason || 'Temporary or disposable burner mailbox'}
                              >
                                <svg className="w-3.5 h-3.5 text-amber-500 flex-shrink-0" viewBox="0 0 20 20" fill="currentColor">
                                  <path fillRule="evenodd" d="M8.485 2.495c.673-1.167 2.357-1.167 3.03 0l6.28 10.875c.673 1.167-.17 2.625-1.516 2.625H3.72c-1.347 0-2.189-1.458-1.515-2.625L8.485 2.495zM10 5a.75.75 0 01.75.75v3.5a.75.75 0 01-1.5 0v-3.5A.75.75 0 0110 5zm0 9a1 1 0 100-2 1 1 0 000 2z" clipRule="evenodd" />
                                </svg>
                                <span>Temp / Disposable</span>
                              </span>
                            ) : mxInfo.accountType === 'consumer' || mxInfo.isFreeMail ? (
                              <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-none text-[10px] font-semibold bg-sky-50 text-sky-700 border border-sky-200 whitespace-nowrap shadow-2xs">
                                <svg className="w-3.5 h-3.5 text-sky-600 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
                                </svg>
                                <span>Personal</span>
                              </span>
                            ) : mxInfo.accountType === 'business' ? (
                              <div className="inline-flex items-center gap-1 whitespace-nowrap">
                                <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-none text-[10px] font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200 shadow-2xs">
                                  <svg className="w-3.5 h-3.5 text-indigo-600 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4" />
                                  </svg>
                                  <span>Corporate</span>
                                </span>
                                {mxInfo.isRoleAccount && (
                                  <span className="px-1.5 py-0.5 rounded-none text-[9px] font-medium bg-purple-50 text-purple-700 border border-purple-200 shadow-2xs" title="Generic role inbox (info@, sales@, etc.)">
                                    Role
                                  </span>
                                )}
                              </div>
                            ) : !mxInfo.hasMx ? (
                              <span className="text-[11px] text-red-500 font-medium whitespace-nowrap">Unreachable</span>
                            ) : (
                              <span className="inline-flex items-center px-2 py-0.5 rounded-none text-[10px] font-medium bg-neutral-100 text-ink-muted whitespace-nowrap">
                                Standard
                              </span>
                            )
                          ) : (
                            <span className="text-[11px] text-ink-muted whitespace-nowrap">Not verified</span>
                          )}
                        </td>

                      </tr>
                    );
                  })}

                </tbody>
              </table>
            </div>
          </div>
          {recipients.length > 10 && (
            <p className="text-xs text-ink-muted text-center">
              Showing preview of first 10 rows out of {recipients.length} total recipients
            </p>
          )}
        </div>
      )}

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
          disabled={recipients.length === 0}
          className="btn btn-primary text-xs px-6 py-2.5 disabled:opacity-40 disabled:cursor-not-allowed shadow-xs"
        >
          Continue to Compose
        </button>
      </div>
    </div>
  );
};

export default RecipientUploader;