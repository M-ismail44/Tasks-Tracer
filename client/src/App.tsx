/** تصميم هذا الملف: «دفتر الأداء التنفيذي» — التطبيق صفحة تشغيل واحدة عربية أولاً. */
import { Toaster } from "@/components/ui/sonner";
import ErrorBoundary from "./components/ErrorBoundary";
import { ThemeProvider } from "./contexts/ThemeContext";
import Home from "./pages/Home";

function App() {
  return <ErrorBoundary><ThemeProvider defaultTheme="light"><Home /><Toaster richColors position="top-left" dir="rtl" /></ThemeProvider></ErrorBoundary>;
}

export default App;
