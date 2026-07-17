import { Sidebar } from './Sidebar';
import { StatusBar } from './StatusBar';
import { PreviewPanel } from '@/components/preview/PreviewPanel';
import { HomePage } from '@/pages/HomePage';
import { PresetsPage } from '@/pages/PresetsPage';
import { ThemePage } from '@/pages/ThemePage';
import { LogPage } from '@/pages/LogPage';
import { useAppStore } from '@/lib/store';
import { AnimatePresence, motion } from 'framer-motion';

const pages = {
  home: HomePage,
  presets: PresetsPage,
  theme: ThemePage,
  log: LogPage,
};

export function Layout() {
  const { activePage } = useAppStore();
  const Page = pages[activePage];

  return (
    <div className="flex h-screen w-screen overflow-hidden">
      <Sidebar />
      <main className="flex-1 flex flex-col min-w-0">
        <div className="flex-1 overflow-hidden relative">
          <AnimatePresence mode="wait">
            <motion.div
              key={activePage}
              initial={{ opacity: 0, y: 6 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -6 }}
              transition={{ duration: 0.15 }}
              className="h-full overflow-auto scrollbar-thin"
            >
              <Page />
            </motion.div>
          </AnimatePresence>
        </div>
        <StatusBar />
      </main>
      <PreviewPanel />
    </div>
  );
}
