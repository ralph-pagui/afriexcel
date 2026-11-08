import { createClient } from '@supabase/supabase-js';

// Configuration Supabase. Ces valeurs peuvent être surchargées via des variables d'environnement
// ou directement ici. Par défaut, nous configurons des valeurs prêtes à l'emploi.
const supabaseUrl = import.meta.env.VITE_SUPABASE_URL || 'https://demo-african-assistant.supabase.co';
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.demo-key';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
