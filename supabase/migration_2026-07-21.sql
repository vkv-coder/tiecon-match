-- ============================================================
-- TIEcon Match-a-Thon: schema + lookup-data migration
-- Run in Supabase SQL Editor. Idempotent / safe to re-run.
-- Covers the July 2026 client review doc changes.
-- ============================================================

-- 1. New columns on tie_experts (startup registration)
ALTER TABLE tie_experts
  ADD COLUMN IF NOT EXISTS linkedin_url text,
  ADD COLUMN IF NOT EXISTS whatsapp text,
  ADD COLUMN IF NOT EXISTS stage text,                  -- 'idea' | 'early-revenue' | 'scaling'
  ADD COLUMN IF NOT EXISTS pitch_deck_url text,
  ADD COLUMN IF NOT EXISTS is_student_led boolean DEFAULT false,
  ADD COLUMN IF NOT EXISTS is_women_led boolean DEFAULT false,
  ADD COLUMN IF NOT EXISTS is_service_provider boolean DEFAULT false;

-- 2. New columns on tie_industries (industry registration)
ALTER TABLE tie_industries
  ADD COLUMN IF NOT EXISTS linkedin_url text,
  ADD COLUMN IF NOT EXISTS whatsapp text,
  ADD COLUMN IF NOT EXISTS timeline text,               -- 'immediate' | 'short-term' | 'long-term'
  ADD COLUMN IF NOT EXISTS budget_range text,
  ADD COLUMN IF NOT EXISTS secondary_industry_type_id uuid REFERENCES tie_industry_types(id);

-- 3. Seed new rows into tie_industry_types ("Industries You Serve")
INSERT INTO tie_industry_types (name, is_active)
SELECT v.name, true
FROM (VALUES
  ('BioTech'),
  ('Aerospace, Aviation & SpaceTech'),
  ('Fashion & Lifestyle'),
  ('Toys & Consumer Products'),
  ('Mining & Industrial Services')
) AS v(name)
WHERE NOT EXISTS (
  SELECT 1 FROM tie_industry_types t WHERE t.name = v.name
);

-- 4. Seed new rows into tie_problem_domains ("Challenge Areas / Industry Type")
INSERT INTO tie_problem_domains (name, is_active)
SELECT v.name, true
FROM (VALUES
  ('Gaming, Sports & Entertainment Tech'),
  ('TravelTech & Tourism'),
  ('PropTech & ConstructionTech'),
  ('EV & Mobility Tech'),
  ('AdTech & MarTech'),
  ('Sector-Agnostic / Horizontal SaaS')
) AS v(name)
WHERE NOT EXISTS (
  SELECT 1 FROM tie_problem_domains t WHERE t.name = v.name
);

-- ============================================================
-- Verify after running:
--   select column_name from information_schema.columns where table_name = 'tie_experts';
--   select column_name from information_schema.columns where table_name = 'tie_industries';
--   select name, is_active from tie_industry_types order by name;
--   select name, is_active from tie_problem_domains order by name;
-- ============================================================
