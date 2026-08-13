import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL =
  import.meta.env.VITE_SUPABASE_URL || "https://yyokwasbsqxwynplbtcm.supabase.co";

const SUPABASE_ANON_KEY =
  import.meta.env.VITE_SUPABASE_ANON_KEY ||
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl5b2t3YXNic3F4d3lucGxidGNtIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzAyMTM2NDEsImV4cCI6MjA4NTc4OTY0MX0.NxwMqD8z7ZTESz3JxcsGf4kbrHvXr7fk3zBfFblmknM";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
