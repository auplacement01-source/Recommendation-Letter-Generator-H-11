# Air University Recommendation Letter Desk

A Streamlit app based on the supplied Air University internship-opportunity letter. It creates individual or Excel-batch letters, assigns a sequential letter number, saves the letter data in Supabase, lets you search/reprint by registration ID, and exports the register to Excel.

## What is included

- `app.py` — Streamlit application.
- `requirements.txt` — Python packages used by the app.
- `air_university_logo.png` — the letterhead logo extracted from the supplied sample; keep this beside `app.py` in GitHub.

The generated Word letters retain the sample's main wording, letterhead, date, student/program/semester sentence, subject line, and signature block. A company-specific letter adds the organization and city to the addressee and names the organization in the request sentence. General letters use the “To Whom it May Concern” salutation.

## Before you start

You will need:

1. A GitHub account.
2. A Streamlit Community Cloud account (sign in with GitHub at [share.streamlit.io](https://share.streamlit.io/)).
3. A Supabase account and project at [supabase.com](https://supabase.com/).

**Privacy note:** student registration IDs and names are personal data. The app uses a shared password and a server-side Supabase key; only share the app with authorized office staff. The shared password is a simple gate, not individual user accounts or an audit trail. Use a strong password and any Streamlit app access restriction available to your account. The Supabase secret/service-role key must never be put in GitHub or shown in a browser.

## Part 1 — Set up the Supabase database

1. Sign in to Supabase and create a project. Choose a strong database password and keep it safe.
2. Open the project dashboard, then go to **SQL Editor** → **New query**.
3. Copy all SQL below into the editor and click **Run**. It creates the letter table, the per-office/program/year counter, and the atomic numbering function.

```sql
create extension if not exists pgcrypto;

create table if not exists public.letter_counters (
    department_prefix text not null,
    program_code text not null,
    year_2d text not null,
    last_number integer not null default 0,
    primary key (department_prefix, program_code, year_2d)
);

create table if not exists public.letters (
    id uuid primary key default gen_random_uuid(),
    letter_number text not null unique,
    student_name text not null,
    registration_id text not null,
    gender text not null,
    program text not null,
    program_code text not null,
    semester smallint not null check (semester between 1 and 12),
    department_prefix text not null,
    year_2d text not null,
    sequence_no integer not null,
    letter_date date not null,
    recipient_mode text not null,
    company_name text not null default '',
    city text not null default '',
    officer_name text not null default 'MOBEEN JAMSHED',
    officer_title text not null default 'Assistant Director',
    office_name text not null default 'Placement & Alumni Affairs',
    created_at timestamptz not null default now()
);

create index if not exists letters_registration_id_idx
    on public.letters (registration_id);
create index if not exists letters_created_at_idx
    on public.letters (created_at desc);

alter table public.letters enable row level security;
alter table public.letter_counters enable row level security;

-- The app uses its key only on the Streamlit server. Do not grant the public
-- anon/authenticated roles access to student records or counter rows.
revoke all on table public.letters from anon, authenticated;
revoke all on table public.letter_counters from anon, authenticated;
grant all on table public.letters to service_role;
grant all on table public.letter_counters to service_role;

create or replace function public.next_letter_serial(
    p_department_prefix text,
    p_program_code text,
    p_year_2d text
)
returns integer
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
    v_serial integer;
begin
    insert into public.letter_counters (department_prefix, program_code, year_2d, last_number)
    values (p_department_prefix, p_program_code, p_year_2d, 1)
    on conflict (department_prefix, program_code, year_2d)
    do update set last_number = public.letter_counters.last_number + 1
    returning last_number into v_serial;

    return v_serial;
end;
$$;

revoke all on function public.next_letter_serial(text, text, text)
    from public, anon, authenticated;
grant execute on function public.next_letter_serial(text, text, text)
    to service_role;
```

**Why the server-side key matters:** the counter function and record operations are intended only for the server. RLS is enabled and public client roles have no table permissions. The app's Supabase key is held in Streamlit Secrets and is never embedded in the browser. Supabase secret/service-role keys bypass RLS, so treat this credential like an administrator password.

## Part 2 — Get the Supabase URL and server-only key

1. In the Supabase project, open **Project Settings** → **API** (the exact menu label can change).
2. Copy the **Project URL**.
3. Copy a server-side key. Use the current **Secret key** if Supabase offers one; otherwise use the legacy `service_role` key. Do **not** use a public/`anon` key for this app, and do not share the server key.

## Part 3 — Put the files in GitHub

1. Create a new GitHub repository, for example `air-university-letter-desk`.
2. Upload `app.py`, `requirements.txt`, and `air_university_logo.png` to the **root** of the repository. You may also upload this README.
3. Do not upload your Supabase key, password, or a `.streamlit/secrets.toml` file. Secrets belong in Streamlit's Secrets manager, not GitHub.

## Part 4 — Deploy from GitHub on Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io/) and sign in with GitHub.
2. Select **Create app** → **Yup, I have an app**.
3. Select your GitHub repository, the branch (usually `main`), and main file path `app.py`.
4. Open **Advanced settings**. Choose a supported Python version; Python 3.12 is a suitable current choice.
5. In the **Secrets** field, paste this TOML and replace the example values:

   ```toml
   [app]
   password = "REPLACE-WITH-A-LONG-UNIQUE-RANDOM-PASSWORD"

   [supabase]
   url = "https://YOUR-PROJECT-REF.supabase.co"
   key = "YOUR-SUPABASE-SERVER-ONLY-SECRET-KEY"
   ```

6. Click **Save** / **Deploy** and wait for Streamlit to install the packages and start the app. Choose an app subdomain if offered.
7. When it opens, enter the password you set in `[app]` to unlock the workspace. Test with one non-sensitive test record first.

Streamlit may label deployment settings differently over time. If you need to add or update secrets later, open the app's settings in your Streamlit workspace and edit **Secrets**. Never commit secret values to GitHub.

## Part 5 — Use the app

### Make one letter

1. Choose the office/department prefix in the left sidebar.
2. Enter the student's full name and registration ID.
3. Choose the program and semester.
4. Review the suggested pronoun and correct it if needed.
5. Choose **General** or **Company / organization**. For a company letter, enter the organization; city is optional.
6. Click **Generate and save letter**. The app saves a record in Supabase and provides a Word `.docx` to download or print.

### Make letters from an Excel roster

1. In **Bulk from Excel**, download the blank template.
2. Fill in `student_name` and `registration_id` for each person. Program and semester can be left blank to use the selections in the app. Gender can be `Female`, `Male`, or `They`; leave it blank for a best-effort guess. For company-targeted letters, set `recipient_mode` to `Company / organization` and fill `company_name` (optionally `city`).
3. Upload the `.xlsx` file (CSV is also accepted), set defaults, review import notes, and generate.
4. Download the single combined Word document (one page per valid student) and/or the batch Excel register.

### Find and reprint a lost letter

Open **Find / reprint**, enter the student's exact registration ID, search, then download the saved letter again. The letter number and stored details are reused.

### Export the database register

Open **Records & export** → **Load saved records**. Filter if needed, then download the visible records as Excel. Supabase remains the primary record store; keep a separate office backup of exports.

## Letter numbering

Default format is the requested pattern:

`IBD/AU/DSA/PLAC/H-11/BBA/26/001`

The number uses the selected office prefix, the program code, the two-digit year from the letter date, and a sequence allocated in Supabase. The sequence increments separately for each prefix + program + year, including when more than one staff member generates letters at the same time. It restarts at `001` for a new year or different office/program combination. Gaps can occur if a number is allocated but the subsequent database insert fails; numbers are not reused. The optional “Include semester in letter number” setting adds an `S3`-style segment.

The supplied document used a different historical numbering style (`.../26/PLAC/H-11/MSCP/S4/091`). The app defaults to your written example, but the office prefix is selectable/customizable in the sidebar and the semester segment can be enabled. Confirm your office's final numbering convention before issuing official letters.

## Important pronoun limitation

The automatic gender guess is an offline, small name-list heuristic. Pakistani names can be unisex, spelled many ways, and used differently across families; a name alone cannot reliably establish gender. Unknown names default to neutral **They** and are flagged in the single-letter screen. Review the selector for every letter. The Excel `gender` column is an explicit override.

## Troubleshooting

- **“Supabase is not configured”** — verify `[supabase] url` and `[supabase] key` are in the app's Streamlit Secrets, then reboot the app.
- **RPC / `next_letter_serial` error** — rerun the SQL setup in Part 1 and confirm the app uses the server-only Supabase key.
- **Permission / RLS errors** — do not solve these by enabling public access. Verify the key is a server-side secret/service-role key and the SQL grants/function permissions ran successfully.
- **Logo missing** — ensure `air_university_logo.png` is in the same GitHub folder as `app.py`.
- **Wrong letter number prefix** — change the office prefix in the sidebar before generating. Existing records keep their original numbers.
- **Excel rows skipped** — expand the import notes; check spelling of required column headers and company fields for company-specific rows.

## Official reference pages

- [Deploy an app on Streamlit Community Cloud](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)
- [Streamlit Community Cloud secrets management](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management)
- [Supabase Row Level Security](https://supabase.com/docs/guides/database/postgres/row-level-security)
- [Supabase API keys](https://supabase.com/docs/guides/getting-started/api-keys)
