import { supabase } from './supabaseClient';

export async function fetchPersonale(commandId) {
  const { data, error } = await supabase
    .from('users')
    .select('id, name, surname, matricola, role, command_id, active, email')
    .eq('command_id', commandId)
    .eq('active', true)
    .order('surname', { ascending: true });

  if (error) {
    console.error('Errore caricamento personale:', error);
    return [];
  }

  return data || [];
}
