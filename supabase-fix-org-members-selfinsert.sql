-- Rotary_Events project (shared by tiecon-match/tie_, clubs, events,
-- registrations, and any other app built on the generic
-- organizations/org_members layer) — fix cross-tenant privilege
-- escalation via org_members self-insert.
--
-- Same bug class as Derasar Boli's dr_profiles and Reminders'
-- rm_users, both already fixed. "members_insert_self" only checked
-- (user_id = auth.uid()) - organization_id and role were fully
-- client-controlled, and is_org_member() (the trust function every
-- other policy here relies on - organizations, org_modules, and
-- presumably tie_*'s own tenant scoping) has no approval concept to
-- catch it. Any authenticated user could self-insert
-- {organization_id: <any existing org>, role: 'admin'} and get
-- immediate access to that org's modules/data.
--
-- Fix: same pattern as Reminders - self-insert only allowed as the
-- FIRST member of a brand-new org (a fresh organization_id has zero
-- existing org_members rows). Joining an org that already has a
-- member needs an explicit invite path (none currently exists in this
-- schema via RLS - worth adding before this system goes live with
-- real invite flows, flagged as a follow-up, not fixed here).
--
-- Verified safe before applying: organizations and org_members are
-- both currently EMPTY (this system isn't live with real data yet),
-- so this is a genuine no-op today - the safest possible time to close
-- this before real orgs start using it.

drop policy if exists "members_insert_self" on org_members;

create policy "members_insert_self" on org_members
  for insert to public
  with check (
    user_id = auth.uid()
    and not exists (select 1 from org_members m2 where m2.organization_id = org_members.organization_id)
  );
