import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
const PROJECT_ID = import.meta.env.VITE_PROJECT_ID || 'default';
const TABLE = 'remont_data';

if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
  console.error(
    'Не заданы VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY. См. README.md — нужно создать .env.local (для локальной разработки) и переменные окружения в Vercel (для деплоя).'
  );
}

export const supabase = SUPABASE_URL && SUPABASE_ANON_KEY ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

// Identifies this browser tab so we can ignore realtime echoes of our own writes.
const CLIENT_ID = Math.random().toString(36).slice(2) + Date.now().toString(36);

export async function loadSharedState(defaultData) {
  if (!supabase) return defaultData;

  const { data: row, error } = await supabase.from(TABLE).select('data').eq('id', PROJECT_ID).maybeSingle();

  if (error) {
    console.error('Ошибка загрузки данных из Supabase', error);
    return defaultData;
  }

  if (row && row.data) return row.data;

  // No row yet for this project — seed it with the defaults so everyone starts from the same state.
  const { error: insertError } = await supabase
    .from(TABLE)
    .insert({ id: PROJECT_ID, data: defaultData, updated_by: CLIENT_ID });
  if (insertError) console.error('Ошибка создания начальной записи в Supabase', insertError);

  return defaultData;
}

export async function saveSharedState(data) {
  if (!supabase) return;
  const { error } = await supabase
    .from(TABLE)
    .upsert({ id: PROJECT_ID, data, updated_by: CLIENT_ID, updated_at: new Date().toISOString() });
  if (error) console.error('Ошибка сохранения данных в Supabase', error);
}

// Calls onRemoteChange(data) whenever another browser/tab updates the shared row.
// Returns an unsubscribe function.
export function subscribeToSharedState(onRemoteChange) {
  if (!supabase) return () => {};

  const channel = supabase
    .channel(`remont_data_${PROJECT_ID}`)
    .on(
      'postgres_changes',
      { event: '*', schema: 'public', table: TABLE, filter: `id=eq.${PROJECT_ID}` },
      (payload) => {
        const row = payload.new;
        if (!row || row.updated_by === CLIENT_ID) return; // ignore our own writes
        onRemoteChange(row.data);
      }
    )
    .subscribe();

  return () => {
    supabase.removeChannel(channel);
  };
}
