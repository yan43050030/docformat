import { Layout } from '@/components/layout/Layout';
import { ToastProvider } from '@/components/ui/toast';

export default function App() {
  return (
    <ToastProvider>
      <Layout />
    </ToastProvider>
  );
}
