import { supabase } from "./supabaseClient";

// Este adaptador imita a mesma API do "window.storage" usada durante o
// desenvolvimento no Claude, só que agora gravando de verdade no Postgres
// do Supabase. Assim o App.jsx não precisou ser reescrito.
//
// Tabela esperada (ver supabase-schema.sql):
//   kv_store(key text primary key, value text not null, updated_at timestamptz)

const TABLE = "kv_store";

export const supabaseStorageAdapter = {
  async get(key /*, shared */) {
    const { data, error } = await supabase.from(TABLE).select("value").eq("key", key).maybeSingle();
    if (error) throw error;
    if (!data) throw new Error(`Chave "${key}" não encontrada`);
    return { key, value: data.value, shared: true };
  },

  async set(key, value /*, shared */) {
    const { error } = await supabase.from(TABLE).upsert({ key, value, updated_at: new Date().toISOString() });
    if (error) throw error;
    return { key, value, shared: true };
  },

  async delete(key /*, shared */) {
    const { error } = await supabase.from(TABLE).delete().eq("key", key);
    if (error) throw error;
    return { key, deleted: true, shared: true };
  },

  async list(prefix = "" /*, shared */) {
    let query = supabase.from(TABLE).select("key");
    if (prefix) query = query.like("key", `${prefix}%`);
    const { data, error } = await query;
    if (error) throw error;
    return { keys: (data || []).map((r) => r.key), prefix, shared: true };
  },
};
