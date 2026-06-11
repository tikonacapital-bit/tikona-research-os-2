-- Create user_notifications table
CREATE TABLE IF NOT EXISTS public.user_notifications (
  id uuid NOT NULL DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL,
  title text NOT NULL,
  body text NOT NULL,
  data jsonb NOT NULL DEFAULT '{}'::jsonb,
  read boolean NOT NULL DEFAULT false,
  created_at timestamp with time zone NOT NULL DEFAULT timezone('utc'::text, now()),
  CONSTRAINT user_notifications_pkey PRIMARY KEY (id),
  CONSTRAINT user_notifications_user_id_fkey FOREIGN KEY (user_id) REFERENCES auth.users(id) ON DELETE CASCADE
);

-- Enable Row Level Security (RLS)
ALTER TABLE public.user_notifications ENABLE ROW LEVEL SECURITY;

-- Drop existing policies if they exist (to prevent errors during reapplies)
DROP POLICY IF EXISTS "Users can view their own notifications" ON public.user_notifications;
DROP POLICY IF EXISTS "Users can update their own notifications" ON public.user_notifications;
DROP POLICY IF EXISTS "Service role can do everything on notifications" ON public.user_notifications;

-- Create policies
CREATE POLICY "Users can view their own notifications" ON public.user_notifications
  FOR SELECT TO authenticated USING (auth.uid() = user_id);

CREATE POLICY "Users can update their own notifications" ON public.user_notifications
  FOR UPDATE TO authenticated USING (auth.uid() = user_id);

CREATE POLICY "Service role can do everything on notifications" ON public.user_notifications
  FOR ALL TO service_role USING (true);
