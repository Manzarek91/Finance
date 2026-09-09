import React, { useEffect, useState } from "react";
import { supabase } from "./supabaseClient.js";
import { setStorageUser } from "./storageShim.js";

const LOCAL_PREFIX = "libro-mastro:";

function hasLocalData() {
  for (let i = 0; i < localStorage.length; i++) {
    if (localStorage.key(i)?.startsWith(LOCAL_PREFIX)) return true;
  }
  return false;
}

async function migrateLocalData(userId) {
  const entries = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (k && k.startsWith(LOCAL_PREFIX)) {
      entries.push({
        user_id: userId,
        key: k.slice(LOCAL_PREFIX.length),
        value: localStorage.getItem(k),
        updated_at: new Date().toISOString(),
      });
    }
  }
  if (!entries.length) return 0;
  const { error } = await supabase
    .from("kv_store")
    .upsert(entries, { onConflict: "user_id,key" });
  if (error) throw error;
  return entries.length;
}

export default function AuthGate({ children }) {
  const [session, setSession] = useState(undefined); // undefined = ancora in caricamento
  const [email, setEmail] = useState("");
  const [sent, setSent] = useState(false);
  const [error, setError] = useState("");
  const [migrated, setMigrated] = useState(false);
  const [migrating, setMigrating] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => listener.subscription.unsubscribe();
  }, []);

  // Impostato direttamente nel corpo del componente (non in un effect) così
  // è già pronto prima che i figli (il portale) montino e leggano i dati.
  if (session && session.user) {
    setStorageUser(session.user.id);
  }

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

  const runMigration = async () => {
    setMigrating(true);
    setError("");
    try {
      await migrateLocalData(session.user.id);
      setMigrated(true);
    } catch (e) {
      setError("Migrazione non riuscita: " + e.message);
    } finally {
      setMigrating(false);
    }
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

  return (
    <>
      {hasLocalData() && !migrated && (
        <div
          style={{
            background: "#FFF7E6",
            padding: "10px 16px",
            fontSize: 13,
            display: "flex",
            alignItems: "center",
            gap: 10,
            justifyContent: "center",
            flexWrap: "wrap",
            color: "#6B5B33",
          }}
        >
          Ho trovato dati salvati in questo browser da prima del login.
          <button
            onClick={runMigration}
            disabled={migrating}
            style={{
              padding: "4px 10px",
              borderRadius: 6,
              border: "1px solid #C6924A",
              background: "#fff",
              cursor: "pointer",
              fontSize: 13,
            }}
          >
            {migrating ? "Importazione…" : "Importali nel tuo account"}
          </button>
        </div>
      )}
      {migrated && (
        <div style={{ background: "#E4F1EA", padding: "8px 16px", fontSize: 13, textAlign: "center", color: "#2E8F6F" }}>
          Dati importati! Ricarica la pagina per vederli.
        </div>
      )}
      {error && (
        <div style={{ background: "#F6E5E1", padding: "8px 16px", fontSize: 13, textAlign: "center", color: "#C2513F" }}>
          {error}
        </div>
      )}
      {children}
    </>
  );
}
