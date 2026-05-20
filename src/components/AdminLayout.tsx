import {
  Building2,
  Globe,
  FileText,
  BookOpen,
  LayoutDashboard,
  GitBranch,
  Send,
  Layers,
  Activity,
} from 'lucide-react';
import SidebarLayout, { type NavItemConfig } from '@/components/Layout';

const adminNavItems: NavItemConfig[] = [
  { label: 'Dashboard',         path: '/admin',                   icon: LayoutDashboard, end: true },
  { label: 'Equity Database',   path: '/admin/equity-database',   icon: Building2,       group: 'Data' },
  { label: 'Equity Universe',   path: '/admin/universe',          icon: Globe,           group: 'Data' },
  { label: 'Research Reports',  path: '/admin/research-reports',  icon: FileText,        group: 'Research' },
  { label: 'Report Generator',  path: '/admin/pipeline',          icon: GitBranch,       group: 'Research' },
  { label: 'Sector Thesis',     path: '/admin/sector-thesis',     icon: Layers,          group: 'Research' },
  { label: 'Prompt Library',    path: '/admin/prompts',           icon: BookOpen,        group: 'Research' },
  { label: 'Recommendations',   path: '/admin/recommendations',   icon: Send,            group: 'Distribution' },
  { label: 'System Health',     path: '/admin/system-health',     icon: Activity,        group: 'System' },
];

export default function AdminLayout() {
  return (
    <SidebarLayout
      navItems={adminNavItems}
      logoSubtitle="Research OS"
      logoLink="/admin"
    />
  );
}
