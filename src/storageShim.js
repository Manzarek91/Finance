// Ricrea l'API window.storage (get/set/delete/list) usata dall'app,
// ma appoggiandosi a Supabase invece che al browser: così i dati
// diventano accessibili da qualsiasi dispositivo dopo il login.
// Ogni riga è vincolata al tuo utente tramite Row Level Security su Supabase.

import { supabase } from "./supabaseClient.js";

let currentUserId = null;

export function setStorageUser(userId) {
  currentUserId = userId;
}

function requireUser() {
  if (!currentUserId) {
    throw new Error("Devi effettuare l'accesso prima di caricare o salvare dati.");
  }
  return currentUserId;
}

window.storage = {
  async get(key) {
    const user_id = requireUser();
    const { data, error } = await supabase
      .from("kv_store")
      .select("value")
      .eq("user_id", user_id)
      .eq("key", key)
      .maybeSingle();
    if (error) throw new Error("Errore di lettura da Supabase: " + error.message);
    if (!data) return null;
    return { key, value: data.value, shared: false };
  },

  async set(key, value) {
    const user_id = requireUser();
    const { error } = await supabase
      .from("kv_store")
      .upsert(
        { user_id, key, value, updated_at: new Date().toISOString() },
        { onConflict: "user_id,key" }
      );
    if (error) throw new Error("Errore di scrittura su Supabase: " + error.message);
    return { key, value, shared: false };
  },

  async delete(key) {
    const user_id = requireUser();
    const { error } = await supabase
      .from("kv_store")
      .delete()
      .eq("user_id", user_id)
      .eq("key", key);
    if (error) throw new Error("Errore di cancellazione da Supabase: " + error.message);
    return { key, deleted: true, shared: false };
  },

  async list(prefix = "") {
    const user_id = requireUser();
    const { data, error } = await supabase
      .from("kv_store")
      .select("key")
      .eq("user_id", user_id)
      .like("key", `${prefix}%`);
    if (error) throw new Error("Errore di lettura da Supabase: " + error.message);
    return { keys: (data || []).map((r) => r.key), prefix, shared: false };
  },
};
