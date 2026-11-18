import React, { useEffect, useState } from "react";
import { supabase } from "./supabaseClient.js";

export default function AuthGate({ children }) {
  const [session, setSession] = useState(undefined); // undefined = ancora in caricamento
  const [email, setEmail] = useState("");
  const [sent, setSent] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => listener.subscription.unsubscribe();
  }, []);

  const sendLink = async (e) => {
    e.preventDefault();
    setError("");
    const { error } = await supabase.auth.signInWithOtp({
      email,
      options: { emailRedirectTo: window.location.href },
    });
    if (error) setError(error.message);
    else setSent(true);
  };

  if (session === undefined) {
    return (
      <div style={{ padding: 40, fontFamily: "system-ui", color: "#8B8676" }}>Caricamento…</div>
    );
  }

  if (!session) {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: "#F2EEE4",
          fontFamily: "system-ui",
        }}
      >
        <form
          onSubmit={sendLink}
          style={{
            background: "#fff",
            padding: 32,
            borderRadius: 16,
            boxShadow: "0 6px 20px rgba(0,0,0,0.08)",
            width: 320,
          }}
        >
          <div style={{ fontSize: 20, fontWeight: 600, marginBottom: 4, color: "#1C1B18" }}>
            Libro Mastro
          </div>
          <div style={{ fontSize: 13, color: "#8B8676", marginBottom: 16 }}>
            Accedi con la tua email per vedere i tuoi dati da qualsiasi dispositivo.
          </div>
          {sent ? (
            <div style={{ fontSize: 14, color: "#2E8F6F" }}>
              Controlla la tua casella email: ti abbiamo inviato un link di accesso.
            </div>
          ) : (
            <>
              <input
                type="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="tu@esempio.com"
                style={{
                  width: "100%",
                  padding: "8px 10px",
                  border: "1px solid #E6E1D3",
                  borderRadius: 8,
                  marginBottom: 10,
                  boxSizing: "border-box",
                  fontSize: 14,
                }}
              />
              <button
                type="submit"
                style={{
                  width: "100%",
                  padding: "9px 0",
                  background: "#3FA79E",
                  color: "#fff",
                  border: "none",
                  borderRadius: 8,
                  fontWeight: 500,
                  cursor: "pointer",
                  fontSize: 14,
                }}
              >
                Invia link di accesso
              </button>
            </>
          )}
          {error && <div style={{ fontSize: 12, color: "#C2513F", marginTop: 10 }}>{error}</div>}
        </form>
      </div>
    );
  }

  // I figli leggono i dati applicativi direttamente da Supabase. Non viene
  // letto, importato o sincronizzato alcun dato applicativo dal browser.
  return children;
}
