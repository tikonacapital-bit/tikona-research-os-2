import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { lazy, Suspense } from 'react';
import { Toaster } from '@/components/ui/sonner';
import { AuthProvider } from '@/contexts/AuthContext';
import ProtectedLayout from '@/components/ProtectedLayout';
import AdminLayout from '@/components/AdminLayout';
import Login from '@/pages/Login';

// Admin pages
import Dashboard from '@/pages/Dashboard';
import MasterDatabase from '@/pages/MasterDatabase';
import Universe from '@/pages/Universe';
import ResearchReports from '@/pages/ResearchReports';
import PromptLibrary from '@/pages/PromptLibrary';
import ResearchPipeline from '@/pages/ResearchPipeline';
import Recommendations from '@/pages/Recommendations';
import SectorThesis from '@/pages/SectorThesis';
import SystemHealth from '@/pages/SystemHealth';

// Only load devtools in development
const ReactQueryDevtools = import.meta.env.DEV
  ? lazy(() =>
      import('@tanstack/react-query-devtools').then((mod) => ({
        default: mod.ReactQueryDevtools,
      }))
    )
  : () => null;

// Create a client with optimized defaults for financial data
const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      staleTime: 30000, // 30 seconds
      gcTime: 300000, // 5 minutes (formerly cacheTime)
      retry: 2,
      refetchOnWindowFocus: false, // Prevent unnecessary refetches
    },
    mutations: {
      retry: 1,
    },
  },
});

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <BrowserRouter>
        <AuthProvider>
          <Routes>
            {/* Public Routes */}
            <Route path="/login" element={<Login />} />

            {/* Admin Routes */}
            <Route element={<ProtectedLayout />}>
              <Route element={<AdminLayout />}>
                <Route path="/admin" element={<Dashboard />} />
                <Route path="/admin/equity-database" element={<MasterDatabase />} />
                <Route path="/admin/universe" element={<Universe />} />
                <Route path="/admin/research-reports" element={<ResearchReports />} />
                <Route path="/admin/pipeline" element={<ResearchPipeline />} />
                <Route path="/admin/prompts" element={<PromptLibrary />} />
                <Route path="/admin/recommendations" element={<Recommendations />} />
                <Route path="/admin/sector-thesis" element={<SectorThesis />} />
                <Route path="/admin/system-health" element={<SystemHealth />} />
              </Route>
            </Route>

            {/* Catch-all redirect */}
            <Route path="*" element={<Navigate to="/admin" replace />} />
          </Routes>

          {/* Toast notifications */}
          <Toaster position="top-right" richColors closeButton />
        </AuthProvider>
      </BrowserRouter>

      {/* React Query Devtools - only in development */}
      <Suspense fallback={null}>
        <ReactQueryDevtools initialIsOpen={false} />
      </Suspense>
    </QueryClientProvider>
  );
}

export default App;
