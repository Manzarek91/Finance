import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // In produzione questo significa che i secret GitHub Actions
  // (SUPABASE_URL / SUPABASE_ANON_KEY) non sono impostati correttamente.
  console.error(
    "Configurazione Supabase mancante: controlla VITE_SUPABASE_URL e VITE_SUPABASE_ANON_KEY."
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
