import React, { useEffect, useState } from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import LoginScreen from "./LoginScreen";
import { clearStoredToken, fetchSession, getStoredToken } from "./authClient";
import "./styles.css";

function Root() {
  const [ready, setReady] = useState(false);
  const [allowed, setAllowed] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const token = getStoredToken();
      if (!token) {
        if (!cancelled) {
          setAllowed(false);
          setReady(true);
        }
        return;
      }
      const session = await fetchSession(token);
      if (cancelled) return;
      if (session.ok && session.email) {
        setAllowed(true);
      } else {
        clearStoredToken();
        setAllowed(false);
      }
      setReady(true);
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  if (!ready) {
    return (
      <div className="layout">
        <p className="muted">로딩 중…</p>
      </div>
    );
  }

  return allowed ? (
    <App
      onLogout={() => {
        clearStoredToken();
        setAllowed(false);
      }}
    />
  ) : (
    <LoginScreen onSuccess={() => setAllowed(true)} />
  );
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>,
);
