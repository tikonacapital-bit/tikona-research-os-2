import { useState } from 'react';
import {
  BookOpen,
  Plus,
  Trash2,
  Lock,
  Sparkles,
  Loader2,
  Edit2,
  Save,
  ChevronUp,
  ChevronDown,
} from 'lucide-react';
import { useAuth } from '@/contexts/AuthContext';
import { useConfirm } from '@/contexts/ConfirmContext';
import {
  usePromptTemplates,
  useCreatePromptTemplate,
  useUpdatePromptTemplate,
  useDeletePromptTemplate,
  useReorderPromptTemplates,
} from '@/hooks/usePromptTemplate';
import { addReportSectionColumn, dropReportSectionColumn } from '@/lib/api';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { CardSkeleton } from '@/components/ui/spinner';
import { toast } from 'sonner';
import type { PromptTemplate } from '@/types/database';

export default function PromptLibrary() {
  const { user } = useAuth();
  const confirm = useConfirm();
  const userEmail = user?.email;

  const { data: templates, isLoading } = usePromptTemplates(userEmail);
  const createMutation = useCreatePromptTemplate();
  const updateMutation = useUpdatePromptTemplate();
  const deleteMutation = useDeletePromptTemplate();
  const reorderMutation = useReorderPromptTemplates();

  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [editingTemplate, setEditingTemplate] = useState<PromptTemplate | null>(null);
  const [formData, setFormData] = useState({
    section_key: '',
    title: '',
    heading_prompt: '',
    prompt_text: '',
    search_keywords: '',
  });

  const isEditMode = editingTemplate !== null;

  const openAddDialog = () => {
    setEditingTemplate(null);
    setFormData({
      section_key: '',
      title: '',
      heading_prompt: '',
      prompt_text: '',
      search_keywords: '',
    });
    setIsDialogOpen(true);
  };

  const openEditDialog = (template: PromptTemplate) => {
    setEditingTemplate(template);
    setFormData({
      section_key: template.section_key,
      title: template.title,
      heading_prompt: template.heading_prompt || '',
      prompt_text: template.prompt_text,
      search_keywords: template.search_keywords.join(', '),
    });
    setIsDialogOpen(true);
  };

  const closeDialog = () => {
    setIsDialogOpen(false);
    setEditingTemplate(null);
    setFormData({
      section_key: '',
      title: '',
      heading_prompt: '',
      prompt_text: '',
      search_keywords: '',
    });
  };

  const handleSubmit = async () => {
    if (!formData.section_key.trim() || !formData.title.trim() || !formData.prompt_text.trim()) {
      toast.error('Section name, title and prompt text are required');
      return;
    }

    try {
      const keywords = formData.search_keywords
        .split(',')
        .map((kw) => kw.trim())
        .filter(Boolean);

      // Normalize section_key to snake_case for consistent matching
      const normalizedKey = formData.section_key.trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');

      // The 7 default section keys that already have dedicated columns
      const defaultSectionKeys = [
        'company_background', 'business_model', 'management_analysis',
        'industry_overview', 'industry_tailwinds', 'demand_drivers', 'industry_risks',
      ];
      const isCustomSection = !defaultSectionKeys.includes(normalizedKey);

      if (isEditMode && editingTemplate) {
        // Update existing template (both default and custom)
        await updateMutation.mutateAsync({
          id: editingTemplate.id,
          updates: {
            title: formData.title,
            heading_prompt: formData.heading_prompt.trim() || undefined,
            prompt_text: formData.prompt_text,
            search_keywords: keywords.length > 0 ? keywords : [],
            section_key: normalizedKey,
          },
        });
        toast.success('Prompt updated!');
      } else {
        // For custom sections, create the DB column first
        if (isCustomSection) {
          await addReportSectionColumn(normalizedKey);
        }

        // Create new template
        await createMutation.mutateAsync({
          section_key: normalizedKey,
          title: formData.title,
          heading_prompt: formData.heading_prompt.trim() || undefined,
          prompt_text: formData.prompt_text,
          search_keywords: keywords.length > 0 ? keywords : [],
        });
        toast.success('Custom prompt created!');
      }

      closeDialog();
    } catch (error) {
      console.error('[PromptLibrary] Submit error:', error);
      toast.error(
        error instanceof Error ? error.message : `Failed to ${isEditMode ? 'save' : 'create'} prompt`
      );
    }
  };

  const handleDelete = async (template: PromptTemplate) => {
    const proceed = await confirm({
      title: `Delete prompt "${template.title}"?`,
      description: 'This action cannot be undone. All template details will be permanently deleted.',
      confirmText: 'Delete',
      cancelText: 'Cancel',
      variant: 'destructive',
    });
    if (!proceed) {
      return;
    }

    const defaultSectionKeys = [
      'company_background', 'business_model', 'management_analysis',
      'industry_overview', 'industry_tailwinds', 'demand_drivers', 'industry_risks',
    ];

    try {
      await deleteMutation.mutateAsync(template.id);

      // Drop the cs_ column if it's a custom section and no other templates use this key
      const isCustomSection = !defaultSectionKeys.includes(template.section_key);
      if (isCustomSection) {
        // Check if other templates still use this section_key
        const othersWithSameKey = templates?.filter(
          (t) => t.id !== template.id && t.section_key === template.section_key
        );
        if (!othersWithSameKey || othersWithSameKey.length === 0) {
          await dropReportSectionColumn(template.section_key);
        }
      }

      toast.success('Prompt deleted');
    } catch (error) {
      console.error('[PromptLibrary] Delete error:', error);
      toast.error(
        error instanceof Error ? error.message : 'Failed to delete prompt'
      );
    }
  };

  const handleMoveTemplate = async (templateId: string, direction: 'up' | 'down') => {
    if (!templates) return;
    const currentIdx = templates.findIndex((t) => t.id === templateId);
    if (currentIdx === -1) return;
    const targetIdx = direction === 'up' ? currentIdx - 1 : currentIdx + 1;
    if (targetIdx < 0 || targetIdx >= templates.length) return;

    const current = templates[currentIdx];
    const target = templates[targetIdx];

    try {
      await reorderMutation.mutateAsync([
        { id: current.id, sort_order: target.sort_order },
        { id: target.id, sort_order: current.sort_order },
      ]);
    } catch (error) {
      console.error('[PromptLibrary] Reorder error:', error);
      toast.error('Failed to reorder');
    }
  };

  // Group templates by section_key
  const groupedTemplates = templates?.reduce<Record<string, PromptTemplate[]>>((acc, template) => {
    if (!acc[template.section_key]) {
      acc[template.section_key] = [];
    }
    acc[template.section_key].push(template);
    return acc;
  }, {});

  // Get unique section keys in order
  const sectionKeys = groupedTemplates ? Object.keys(groupedTemplates) : [];

  const isPending = createMutation.isPending || updateMutation.isPending;

  return (
    <div className="flex h-full flex-col">
      {/* Page Header */}
      <header className="border-b border-neutral-200/80 bg-white px-7 py-5">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-lg font-semibold tracking-tight text-neutral-900">
              Prompt Library
            </h1>
            <p className="text-sm text-neutral-500">
              {templates ? `${templates.length} prompts` : 'Manage prompts for report generation'}
            </p>
          </div>

          <Button onClick={openAddDialog} size="sm">
            <Plus className="h-4 w-4 mr-2" />
            Add Custom Prompt
          </Button>
        </div>
      </header>

      {/* Main Content */}
      <div className="flex-1 overflow-auto bg-canvas p-7">
        {isLoading ? (
          <div className="max-w-4xl space-y-4">
            <CardSkeleton />
            <CardSkeleton />
            <CardSkeleton />
          </div>
        ) : !templates || templates.length === 0 ? (
          <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-neutral-200 bg-white py-16">
            <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-full bg-accent-50">
              <BookOpen className="h-7 w-7 text-accent-300" />
            </div>
            <h3 className="mt-4 text-sm font-medium text-neutral-900">
              No prompts found
            </h3>
            <p className="mt-1 text-sm text-neutral-500 max-w-sm mx-auto text-center">
              Run the database migration to add default prompts
            </p>
          </div>
        ) : (
          <div className="max-w-4xl space-y-6">
            {sectionKeys.map((sectionKey) => {
              const sectionTemplates = groupedTemplates?.[sectionKey] || [];
              if (sectionTemplates.length === 0) return null;

              return (
                <div key={sectionKey}>
                  <h2 className="text-base font-semibold text-neutral-800 mb-3 px-1">
                    {sectionKey.replace(/_/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase())}
                  </h2>
                  <div className="space-y-2">
                    {sectionTemplates.map((template) => (
                      <div
                        key={template.id}
                        className="card-premium p-5"
                      >
                        <div className="flex items-start justify-between gap-4">
                          <div className="flex items-start gap-3 flex-1 min-w-0">
                            <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-accent-50">
                              {template.is_default ? (
                                <Lock className="h-4 w-4 text-accent-600" />
                              ) : (
                                <Sparkles className="h-4 w-4 text-accent-600" />
                              )}
                            </div>
                            <div className="flex-1 min-w-0">
                              <div className="flex items-center gap-2">
                                <h3 className="text-base font-semibold text-neutral-900">
                                  {template.title}
                                </h3>
                                {template.is_default && (
                                  <span className="inline-flex items-center rounded-md bg-neutral-100 px-2 py-0.5 text-xs font-medium text-neutral-600">
                                    Default
                                  </span>
                                )}
                              </div>
                              {template.heading_prompt && (
                                <p className="mt-2 text-xs text-blue-600/70 leading-relaxed">
                                  <span className="font-medium text-blue-500">Heading:</span> {template.heading_prompt}
                                </p>
                              )}
                              <div className="mt-2 flex flex-wrap gap-1">
                                {template.search_keywords.slice(0, 8).map((kw) => (
                                  <span
                                    key={kw}
                                    className="inline-flex rounded-md bg-accent-50 px-2 py-0.5 text-xs font-medium text-accent-700"
                                  >
                                    {kw}
                                  </span>
                                ))}
                                {template.search_keywords.length > 8 && (
                                  <span className="text-xs text-neutral-400">
                                    +{template.search_keywords.length - 8} more
                                  </span>
                                )}
                              </div>
                              <p className="mt-3 text-sm text-neutral-600 leading-relaxed whitespace-pre-wrap max-h-48 overflow-y-auto">
                                {template.prompt_text}
                              </p>
                            </div>
                          </div>

                          <div className="flex items-center gap-1">
                            <div className="flex flex-col gap-1 mr-1">
                              <button
                                onClick={() => handleMoveTemplate(template.id, 'up')}
                                disabled={!templates || templates.indexOf(template) === 0 || reorderMutation.isPending}
                                className="h-5 w-5 flex items-center justify-center rounded text-neutral-400 hover:text-neutral-700 hover:bg-neutral-100 disabled:opacity-25 disabled:cursor-not-allowed transition-colors"
                                title="Move up"
                              >
                                <ChevronUp className="h-3.5 w-3.5" />
                              </button>
                              <button
                                onClick={() => handleMoveTemplate(template.id, 'down')}
                                disabled={!templates || templates.indexOf(template) === templates.length - 1 || reorderMutation.isPending}
                                className="h-5 w-5 flex items-center justify-center rounded text-neutral-400 hover:text-neutral-700 hover:bg-neutral-100 disabled:opacity-25 disabled:cursor-not-allowed transition-colors"
                                title="Move down"
                              >
                                <ChevronDown className="h-3.5 w-3.5" />
                              </button>
                            </div>
                            <Button
                              size="sm"
                              variant="ghost"
                              onClick={() => openEditDialog(template)}
                              className="h-8 w-8 p-0"
                              title="Edit prompt"
                            >
                              <Edit2 className="h-4 w-4" />
                            </Button>
                            <Button
                              size="sm"
                              variant="ghost"
                              onClick={() => handleDelete(template)}
                              disabled={deleteMutation.isPending}
                              className="h-8 w-8 p-0 text-red-600 hover:text-red-700 hover:bg-red-50"
                              title="Delete prompt"
                            >
                              {deleteMutation.isPending ? (
                                <Loader2 className="h-4 w-4 animate-spin" />
                              ) : (
                                <Trash2 className="h-4 w-4" />
                              )}
                            </Button>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Add/Edit Prompt Dialog */}
      <Dialog open={isDialogOpen} onOpenChange={(open) => !open && closeDialog()}>
        <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>
              {isEditMode ? 'Edit Prompt' : 'Add Custom Prompt'}
            </DialogTitle>
            <DialogDescription>
              {isEditMode
                ? 'Update this prompt template'
                : 'Create a custom prompt template for report generation'}
            </DialogDescription>
          </DialogHeader>

          <div className="space-y-4 py-4">
            <div>
              <Label htmlFor="section_key">Section Name</Label>
              <Input
                id="section_key"
                value={formData.section_key}
                onChange={(e) =>
                  setFormData({
                    ...formData,
                    section_key: e.target.value,
                  })
                }
                placeholder="e.g., Company Overview, Financial Summary"
                className="mt-2"
              />
              <p className="mt-1 text-xs text-neutral-500">
                Section name (e.g. &quot;Valuation Analysis&quot;). Will be normalized to snake_case automatically.
              </p>
            </div>

            <div>
              <Label htmlFor="title">Prompt Title</Label>
              <Input
                id="title"
                value={formData.title}
                onChange={(e) =>
                  setFormData({ ...formData, title: e.target.value })
                }
                placeholder="e.g., Detailed Company Overview"
                className="mt-2"
              />
            </div>

            <div>
              <Label htmlFor="heading_prompt">Heading Prompt</Label>
              <textarea
                id="heading_prompt"
                value={formData.heading_prompt}
                onChange={(e) =>
                  setFormData({ ...formData, heading_prompt: e.target.value })
                }
                placeholder="e.g., Create a compelling heading that captures the company's core identity and business essence..."
                rows={3}
                className="mt-2 w-full rounded-lg border border-neutral-200 bg-white px-3 py-2 text-sm focus-visible:border-accent-400 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent-500/40 resize-y leading-relaxed"
              />
              <p className="mt-1 text-xs text-neutral-500">
                Instructions for AI to generate a dynamic heading for this section
              </p>
            </div>

            <div>
              <Label htmlFor="search_keywords">
                Search Keywords (comma-separated)
              </Label>
              <Input
                id="search_keywords"
                value={formData.search_keywords}
                onChange={(e) =>
                  setFormData({ ...formData, search_keywords: e.target.value })
                }
                placeholder="company, overview, business, products, management"
                className="mt-2"
              />
              <p className="mt-1 text-xs text-neutral-500">
                Keywords used for document retrieval during RAG
              </p>
            </div>

            <div>
              <Label htmlFor="prompt_text">Prompt Text</Label>
              <textarea
                id="prompt_text"
                value={formData.prompt_text}
                onChange={(e) =>
                  setFormData({ ...formData, prompt_text: e.target.value })
                }
                placeholder="Write a comprehensive section about..."
                rows={14}
                className="mt-2 w-full rounded-lg border border-neutral-200 bg-white px-3 py-2 text-sm focus-visible:border-accent-400 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent-500/40 resize-y font-mono leading-relaxed"
              />
            </div>
          </div>

          <DialogFooter>
            <Button
              variant="outline"
              onClick={closeDialog}
              disabled={isPending}
            >
              Cancel
            </Button>
            <Button
              onClick={handleSubmit}
              disabled={isPending}
              className="min-w-[140px]"
            >
              {isPending ? (
                <>
                  <Loader2 className="h-4 w-4 mr-2 animate-spin" />
                  {isEditMode ? 'Saving...' : 'Creating...'}
                </>
              ) : isEditMode ? (
                <>
                  <Save className="h-4 w-4 mr-2" />
                  Save Changes
                </>
              ) : (
                <>
                  <Plus className="h-4 w-4 mr-2" />
                  Create Prompt
                </>
              )}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
