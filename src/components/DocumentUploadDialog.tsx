import { useState, useRef, useCallback } from 'react';
import { Upload, FileText, Loader2, AlertCircle } from 'lucide-react';
import { Button } from '@/components/ui/button';
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
} from '@/components/ui/dialog';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';
import { Label } from '@/components/ui/label';
import { uploadDocument } from '@/lib/api';
import type { VaultDocument } from '@/types/vault';

const DOCUMENT_CATEGORIES = [
  { value: 'annual_report', label: 'Annual Report' },
  { value: 'investor_presentation', label: 'Investor Presentation' },
  { value: 'concall_transcript', label: 'Concall Transcript' },
  { value: 'broker_report', label: 'Broker Report' },
  { value: 'financial_model', label: 'Financial Model' },
  { value: 'other', label: 'Other Document' },
] as const;

const MAX_FILE_SIZE_MB = 50;
const MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024;

const EXCEL_MIME_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
  'application/vnd.ms-excel.sheet.macroEnabled.12',
];

function isExcelFile(file: File): boolean {
  return /\.(xlsx|xls|xlsm)$/i.test(file.name) || EXCEL_MIME_TYPES.includes(file.type);
}

function isAllowedFile(file: File, category: string): boolean {
  if (category === '') {
    return isExcelFile(file) || file.type === 'application/pdf';
  }
  return category === 'financial_model' ? isExcelFile(file) : file.type === 'application/pdf';
}

interface DocumentUploadDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  folderId: string;
  nseSymbol: string;
  onUploadComplete: (document: VaultDocument) => void;
  /** Called after a successful upload when the "Financial Model" category is selected, so the
   *  caller can also mirror the Excel file to the canonical financial-model storage location. */
  onFinancialModelFileUploaded?: (file: File) => void | Promise<void>;
}

export default function DocumentUploadDialog({
  open,
  onOpenChange,
  folderId,
  nseSymbol,
  onUploadComplete,
  onFinancialModelFileUploaded,
}: DocumentUploadDialogProps) {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [category, setCategory] = useState<string>('');
  const [isUploading, setIsUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const resetState = useCallback(() => {
    setSelectedFile(null);
    setCategory('');
    setIsUploading(false);
    setError(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  }, []);

  const handleOpenChange = useCallback(
    (nextOpen: boolean) => {
      if (!nextOpen) {
        resetState();
      }
      onOpenChange(nextOpen);
    },
    [onOpenChange, resetState]
  );

  const handleCategoryChange = useCallback(
    (value: string) => {
      setCategory(value);
      setError(null);
      // Drop the selected file if it no longer matches the newly chosen category's allowed type.
      if (selectedFile && !isAllowedFile(selectedFile, value)) {
        setSelectedFile(null);
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    },
    [selectedFile]
  );

  const handleFileSelect = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      setError(null);

      if (!file) return;

      if (file.size > MAX_FILE_SIZE_BYTES) {
        setError(`File size exceeds ${MAX_FILE_SIZE_MB}MB limit.`);
        return;
      }

      if (!isAllowedFile(file, category)) {
        setError(category === 'financial_model' ? 'Only Excel files (.xlsx/.xls/.xlsm) are supported.' : 'Only PDF files are supported.');
        return;
      }

      setSelectedFile(file);
      if (category === '' && isExcelFile(file)) {
        setCategory('financial_model');
      }
    },
    [category, setCategory]
  );

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      const file = e.dataTransfer.files?.[0];
      setError(null);

      if (!file) return;

      if (file.size > MAX_FILE_SIZE_BYTES) {
        setError(`File size exceeds ${MAX_FILE_SIZE_MB}MB limit.`);
        return;
      }

      if (!isAllowedFile(file, category)) {
        setError(category === 'financial_model' ? 'Only Excel files (.xlsx/.xls/.xlsm) are supported.' : 'Only PDF files are supported.');
        return;
      }

      setSelectedFile(file);
      if (category === '' && isExcelFile(file)) {
        setCategory('financial_model');
      }
    },
    [category, setCategory]
  );

  const handleUpload = useCallback(async () => {
    if (!selectedFile || !category) return;

    setIsUploading(true);
    setError(null);

    try {
      // Read file as base64
      const base64 = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
          const result = reader.result as string;
          // Remove the data:application/pdf;base64, prefix
          const base64Data = result.split(',')[1];
          resolve(base64Data);
        };
        reader.onerror = reject;
        reader.readAsDataURL(selectedFile);
      });

      // Use original filename as-is (no renaming)
      const fileName = selectedFile.name;

      // Map UI category to the standardized subfolder name in Google Drive
      const categoryToSubfolderMap: Record<string, string> = {
        annual_report: 'Annual Report',
        investor_presentation: 'Investor PPT',
        concall_transcript: 'Concall Transcript',
        broker_report: 'Broker Report',
        financial_model: 'Financial Model',
        other: 'Other Document',
      };
      const subfolderName = categoryToSubfolderMap[category] || 'Other Document';

      const uploadedDoc = await uploadDocument(folderId, fileName, base64, subfolderName);

      if (category === 'financial_model' && onFinancialModelFileUploaded) {
        await onFinancialModelFileUploaded(selectedFile);
      }

      onUploadComplete(uploadedDoc);
      handleOpenChange(false);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Upload failed. Please try again.'
      );
    } finally {
      setIsUploading(false);
    }
  }, [selectedFile, category, folderId, nseSymbol, onUploadComplete, onFinancialModelFileUploaded, handleOpenChange]);

  const formatSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  return (
    <Dialog open={open} onOpenChange={handleOpenChange}>
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle>Upload Document</DialogTitle>
          <DialogDescription>
            Upload a document to the research vault.
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-4 py-2">
          {/* Category Selection */}
          <div className="space-y-2">
            <Label htmlFor="category">Document Category</Label>
            <Select value={category} onValueChange={handleCategoryChange}>
              <SelectTrigger id="category">
                <SelectValue placeholder="Select category..." />
              </SelectTrigger>
              <SelectContent>
                {DOCUMENT_CATEGORIES.map((cat) => (
                  <SelectItem key={cat.value} value={cat.value}>
                    {cat.label}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          {/* File Drop Zone */}
          <div className="space-y-2">
            <Label>File</Label>
            <div
              onDrop={handleDrop}
              onDragOver={(e) => e.preventDefault()}
              onClick={() => fileInputRef.current?.click()}
              className="flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed border-neutral-300 px-6 py-8 transition-colors hover:border-neutral-400 hover:bg-neutral-50"
            >
              {selectedFile ? (
                <div className="flex items-center gap-3">
                  <FileText className="h-8 w-8 text-neutral-600" />
                  <div className="text-left">
                    <p className="text-sm font-medium text-neutral-900 truncate max-w-[250px]">
                      {selectedFile.name}
                    </p>
                    <p className="text-xs text-neutral-500">
                      {formatSize(selectedFile.size)}
                    </p>
                  </div>
                </div>
              ) : (
                <>
                  <Upload className="h-8 w-8 text-neutral-400 mb-2" />
                  <p className="text-sm text-neutral-600">
                    Click to select or drag & drop
                  </p>
                  <p className="text-xs text-neutral-400 mt-1">
                    {category === 'financial_model'
                      ? `Excel only (.xlsx/.xls/.xlsm), max ${MAX_FILE_SIZE_MB}MB`
                      : category === ''
                      ? `PDF or Excel, max ${MAX_FILE_SIZE_MB}MB`
                      : `PDF only, max ${MAX_FILE_SIZE_MB}MB`}
                  </p>
                </>
              )}
              <input
                ref={fileInputRef}
                type="file"
                accept={
                  category === 'financial_model'
                    ? '.xlsx,.xls,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.ms-excel.sheet.macroEnabled.12'
                    : category === ''
                    ? '.pdf,application/pdf,.xlsx,.xls,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.ms-excel.sheet.macroEnabled.12'
                    : '.pdf,application/pdf'
                }
                onChange={handleFileSelect}
                className="hidden"
              />
            </div>
          </div>

          {/* Error Message */}
          {error && (
            <div className="flex items-start gap-2 rounded-lg bg-red-50 border border-red-200 p-3">
              <AlertCircle className="h-4 w-4 text-red-600 mt-1 flex-shrink-0" />
              <p className="text-sm text-red-700">{error}</p>
            </div>
          )}
        </div>

        <DialogFooter>
          <Button
            variant="outline"
            onClick={() => handleOpenChange(false)}
            disabled={isUploading}
          >
            Cancel
          </Button>
          <Button
            onClick={handleUpload}
            disabled={!selectedFile || !category || isUploading}
            className="min-w-[120px]"
          >
            {isUploading ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                Uploading...
              </>
            ) : (
              <>
                <Upload className="mr-2 h-4 w-4" />
                Upload
              </>
            )}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
