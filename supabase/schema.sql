-- Schéma de base de données PostgreSQL pour l'Assistant IA Excel Afrique
-- Conçu pour Supabase avec Row Level Security (RLS)

-- 1. Table des profils utilisateurs
CREATE TABLE IF NOT EXISTS public.profiles (
  id UUID REFERENCES auth.users ON DELETE CASCADE PRIMARY KEY,
  email TEXT,
  phone TEXT,
  country_code TEXT DEFAULT 'CI',
  created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 2. Table des abonnements et soldes de crédits
CREATE TABLE IF NOT EXISTS public.subscriptions (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  user_id UUID REFERENCES public.profiles(id) ON DELETE CASCADE NOT NULL,
  plan TEXT CHECK (plan IN ('free', 'pro_mensuel', 'pack_credits')) DEFAULT 'free',
  credits_remaining INT DEFAULT 30, -- 30 crédits offerts à l'inscription
  is_active BOOLEAN DEFAULT true,
  expires_at TIMESTAMP WITH TIME ZONE,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 3. Table des transactions Mobile Money (Wave, Orange Money, MTN MoMo, Moov)
CREATE TABLE IF NOT EXISTS public.mobile_money_transactions (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  user_id UUID REFERENCES public.profiles(id),
  operator TEXT CHECK (operator IN ('WAVE', 'ORANGE_MONEY', 'MTN_MOMO', 'MOOV_MONEY', 'CARD')) NOT NULL,
  phone_number TEXT NOT NULL,
  amount NUMERIC NOT NULL,
  currency TEXT DEFAULT 'FCFA',
  status TEXT CHECK (status IN ('PENDING', 'SUCCESS', 'FAILED')) DEFAULT 'PENDING',
  transaction_ref TEXT UNIQUE NOT NULL,
  gateway_ref TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- 4. Table des logs de tâches exécutées dans Excel
CREATE TABLE IF NOT EXISTS public.task_logs (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  user_id UUID REFERENCES public.profiles(id),
  task_category TEXT CHECK (task_category IN ('chat', 'reconciliation', 'syscohada', 'payroll', 'clean', 'aging', 'formula')),
  prompt_preview TEXT,
  model_used TEXT DEFAULT 'gemini-3.6-flash',
  credits_cost INT DEFAULT 1,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- RLS (Row Level Security)
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.subscriptions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.mobile_money_transactions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.task_logs ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users can view own profile" ON public.profiles FOR SELECT USING (auth.uid() = id);
CREATE POLICY "Users can update own profile" ON public.profiles FOR UPDATE USING (auth.uid() = id);

CREATE POLICY "Users can view own subscription" ON public.subscriptions FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "Users can view own transactions" ON public.mobile_money_transactions FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "Users can view own logs" ON public.task_logs FOR SELECT USING (auth.uid() = user_id);
