import { lazy, Suspense, useRef } from "react";
import AppLayout from "@/components/layout/AppLayout";
import { BrowserRouter, Route, Routes } from "react-router-dom";
import ProtectedRoute from "./components/auth/ProtectedRoute";
import PublicRoute from "./components/auth/PublicRoute";
import ErrorBoundary from "./components/common/ErrorBoundary";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { useAuthStore } from "./store/auth-store.js";
import Login from "./components/auth/Login.js";
import LandingPage from "./components/landing/LandingPage.js";
import { Toaster } from "sonner";
// Code splitting: le pagine vengono caricate solo quando servono
const Register = lazy(() => import("./components/auth/Register"));
const RegistraCopie = lazy(() => import("./components/RegistraCopie"));
const GestioneDocenti = lazy(() => import("./components/gestione-docenti/GestioneDocenti"));
const GestioneUtenze = lazy(() => import("./components/gestione-utenze/GestioneUtenze"));
const VisualizzaRegistrazioni = lazy(() => import("./components/visualizza-registrazioni/VisualizzaRegistrazioni"));
const ProfiloUtente = lazy(() => import("./components/profilo/ProfiloUtente"));
const NotFound = lazy(() => import("./components/common/NotFound"));

export function App() {
  const isInitializing = useAuthStore((state) => state.isInitializing);

  // QueryClient creato una sola volta per l'intera vita del componente
  const queryClientRef = useRef<QueryClient | null>(null);
  if (!queryClientRef.current) {
    queryClientRef.current = new QueryClient({
      defaultOptions: {
        queries: {
          retry: 1,
          refetchOnWindowFocus: false,
          staleTime: 2 * 60 * 1000,
          gcTime: 5 * 60 * 1000,
        },
      },
    });
  }

  if (isInitializing) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="h-8 w-8 animate-spin rounded-full border-4 border-primary border-t-transparent" />
      </div>
    );
  }

  const PageFallback = () => (
    <div className="min-h-[50vh] flex items-center justify-center">
      <div className="h-8 w-8 animate-spin rounded-full border-4 border-primary border-t-transparent" />
    </div>
  );

  return (
    <ErrorBoundary>
      <QueryClientProvider client={queryClientRef.current}>
        <BrowserRouter>
          <Suspense fallback={<PageFallback />}>
            <Routes>
            {/* Route pubbliche */}
            <Route path="/" element={<LandingPage />} />
            <Route
              path="/login"
              element={
                <PublicRoute>
                  <Login />
                </PublicRoute>
              }
            />
            <Route
              path="/register"
              element={
                <PublicRoute>
                  <Register />
                </PublicRoute>
              }
            />

            {/* Route protette con AppLayout */}
            <Route
              element={
                <ProtectedRoute>
                  <AppLayout />
                </ProtectedRoute>
              }
            >
              <Route path="/registra-copie" element={<RegistraCopie />} />
              
              <Route
                path="/gestione-docenti"
                element={
                  <ProtectedRoute requiredRole="admin">
                    <GestioneDocenti />
                  </ProtectedRoute>
                }
              />
              <Route
                path="/gestione-utenze"
                element={
                  <ProtectedRoute requiredRole="admin">
                    <GestioneUtenze />
                  </ProtectedRoute>
                }
              />
              <Route
                path="/visualizza-registrazioni"
                element={
                  <ProtectedRoute requiredRole="admin">
                    <VisualizzaRegistrazioni />
                  </ProtectedRoute>
                }
              />
              <Route path="/profilo" element={<ProfiloUtente />} />
            </Route>

            {/* Route catch-all per 404 */}
            <Route path="*" element={<NotFound />} />
          </Routes>
          </Suspense>
        </BrowserRouter>
      </QueryClientProvider>
      <Toaster position="bottom-right"/>
    </ErrorBoundary>
  );
}

export default App;
