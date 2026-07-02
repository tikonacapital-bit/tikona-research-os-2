import * as React from 'react';
import {
  AlertDialog,
  AlertDialogContent,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogCancel,
  AlertDialogAction,
} from '@/components/ui/alert-dialog';
import { AlertTriangle, HelpCircle, Info } from 'lucide-react';
import { cn } from '@/lib/utils';

interface ConfirmOptions {
  title: string;
  description: string;
  confirmText?: string;
  cancelText?: string;
  variant?: 'default' | 'destructive' | 'info';
}

interface ConfirmContextType {
  confirm: (options: ConfirmOptions) => Promise<boolean>;
}

const ConfirmContext = React.createContext<ConfirmContextType | undefined>(undefined);

export function ConfirmProvider({ children }: { children: React.ReactNode }) {
  const [isOpen, setIsOpen] = React.useState(false);
  const [options, setOptions] = React.useState<ConfirmOptions | null>(null);
  const resolverRef = React.useRef<((value: boolean) => void) | null>(null);

  const confirm = React.useCallback((opts: ConfirmOptions) => {
    setOptions(opts);
    setIsOpen(true);
    return new Promise<boolean>((resolve) => {
      resolverRef.current = resolve;
    });
  }, []);

  const handleCancel = React.useCallback(() => {
    setIsOpen(false);
    if (resolverRef.current) {
      resolverRef.current(false);
      resolverRef.current = null;
    }
  }, []);

  const handleConfirm = React.useCallback(() => {
    setIsOpen(false);
    if (resolverRef.current) {
      resolverRef.current(true);
      resolverRef.current = null;
    }
  }, []);

  const variant = options?.variant || 'default';

  return (
    <ConfirmContext.Provider value={{ confirm }}>
      {children}
      <AlertDialog open={isOpen} onOpenChange={(open) => { if (!open) handleCancel(); }}>
        <AlertDialogContent className="max-w-[400px] gap-0 p-6 overflow-hidden">
          <div className="flex gap-4">
            <div className={cn(
              "flex h-10 w-10 shrink-0 items-center justify-center rounded-full",
              variant === 'destructive' && "bg-red-50 text-red-600",
              variant === 'default' && "bg-accent-50 text-accent-600",
              variant === 'info' && "bg-blue-50 text-blue-600"
            )}>
              {variant === 'destructive' && <AlertTriangle className="h-5 w-5" />}
              {variant === 'default' && <HelpCircle className="h-5 w-5" />}
              {variant === 'info' && <Info className="h-5 w-5" />}
            </div>
            <div className="flex-1 space-y-1.5">
              <AlertDialogHeader>
                <AlertDialogTitle className="text-left text-base font-semibold text-neutral-900 leading-snug">
                  {options?.title}
                </AlertDialogTitle>
                <AlertDialogDescription className="text-left text-sm text-neutral-500 leading-relaxed">
                  {options?.description}
                </AlertDialogDescription>
              </AlertDialogHeader>
            </div>
          </div>
          <AlertDialogFooter className="mt-6 gap-2 sm:gap-0">
            <AlertDialogCancel
              onClick={handleCancel}
              className="w-full sm:w-auto h-9 px-4 py-2 text-sm font-medium border border-neutral-200 hover:bg-neutral-50 text-neutral-700 rounded-lg active:scale-[0.98] transition-all"
            >
              {options?.cancelText || 'Cancel'}
            </AlertDialogCancel>
            <AlertDialogAction
              onClick={handleConfirm}
              className={cn(
                "w-full sm:w-auto h-9 px-4 py-2 text-sm font-medium text-white rounded-lg active:scale-[0.98] transition-all",
                variant === 'destructive'
                  ? "bg-red-600 hover:bg-red-700 active:bg-red-800"
                  : "bg-accent-600 hover:bg-accent-700 active:bg-accent-800"
              )}
            >
              {options?.confirmText || 'Confirm'}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </ConfirmContext.Provider>
  );
}

export function useConfirm() {
  const context = React.useContext(ConfirmContext);
  if (!context) {
    throw new Error('useConfirm must be used within a ConfirmProvider');
  }
  return context.confirm;
}
