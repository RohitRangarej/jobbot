# jobbot

A small, self-hosted job finder. It checks company career pages, job boards and your job-alert emails every hour, scores each opening against a profile you write, and logs the good matches to an Excel workbook you can work from: what to apply to today, what to follow up on, and everything found so far, day by day.

- **One Python file, standard library only.** Nothing to install beyond Python 3.9+.
- **Reads, never applies.** It reads public job listings and, if you connect it, your job-alert emails (read-only: it never changes them or marks them as read). It never submits applications.
- **Your data stays on your machine.** Results live in local `.xlsx` files; your email app password lives in macOS Keychain or an environment variable.

## What you get

Each region you define (say "India" and "Singapore") gets its own workbook with these sheets:

| Sheet | What's on it |
|---|---|
| **Apply today** | The 15 best open roles you haven't acted on, ranked by fit, a bonus for companies on your target list, a bonus for good stated pay, and a bonus for postings from the last week |
| **Follow up** | Roles you marked *Applied* 7 or more days ago |
| **Stretch roles** | Good matches that ask for a bit more experience than you have |
| **YYYY-MM-DD** | One sheet per day, with every match found that day |
| **Job boards** | Your own list of links to check by hand |

Every row has the company, role, location, a 0–100 **Fit** score, the experience asked for, the salary when stated, **why it matched** (which of your skills appear), the apply link, and **Status** and **Notes** columns. Status is a dropdown: *To apply, Applied, Interview, Offer, Rejected, Not interested*.

**Your edits are kept.** The bot rebuilds each workbook on every pass but reads back whatever you typed in Status and Notes first. Jobs that disappear from a company's listing are marked `(closed)`.

On macOS you also get a notification when a strong new match appears.

## Quick start

```bash
git clone https://github.com/<your-username>/jobbot.git
cd jobbot
python3 jobbot.py once      # first run creates config.json from config.example.json
```

Then open `config.json`, describe the jobs you want (see [Configuration](#configuration)), and run it again. When you're happy with what it finds, leave it running in the background:

```bash
python3 jobbot.py start     # runs every poll_minutes until stopped; logs to jobbot.log
python3 jobbot.py status    # is it running, and how many matches per region
python3 jobbot.py stop
```

Save the workbook after changing Status or Notes; the next pass reads your edits back in. Reopen it to see newly found jobs.

## Commands

| Command | What it does |
|---|---|
| `python3 jobbot.py once` | One pass in the foreground |
| `python3 jobbot.py start` | Keep running in the background |
| `python3 jobbot.py stop` | Stop the background run |
| `python3 jobbot.py status` | Show whether it's running and what it has found |
| `python3 jobbot.py setup-email` | Connect the Gmail inbox that receives your job alerts |
| `python3 jobbot.py email-check` | Show what it can read from those alerts, without saving anything |

## Where jobs come from

**Company career pages**, through the public job APIs of the applicant tracking systems most tech companies use. You add companies to `feeds` in `config.json`:

| `ats` | `slug` is… | Find it in the careers URL |
|---|---|---|
| `greenhouse` | the board name | `boards.greenhouse.io/<slug>` or `job-boards.greenhouse.io/<slug>` |
| `lever` | the company name | `jobs.lever.co/<slug>` |
| `ashby` | the board name | `jobs.ashbyhq.com/<slug>` |
| `smartrecruiters` | the company id | `jobs.smartrecruiters.com/<slug>` |
| `workable` | the account name | `apply.workable.com/<slug>` |

**Job boards with public listings:**

| `ats` | Covers | `slug` is… | Extra options |
|---|---|---|---|
| `himalayas` | Remote jobs worldwide | a search term, e.g. `backend` | `country`: only roles open to that country. `seniority`: levels to keep (default `["Entry-level", "Mid-level"]`; `[]` keeps all) |
| `cutshort` | India | a listing page, e.g. `backend-developer-jobs-in-bangalore-bengaluru` | |
| `kalibrr` | Indonesia, Philippines | a city, e.g. `Jakarta` | `country` (default `Indonesia`) |
| `dealls` | Indonesia | a city id from dealls.com | |

**Your job-alert emails** (optional). Set up alerts on LinkedIn, Indeed, Naukri and the rest, connect the Gmail inbox they go to, and the bot reads the job cards out of them. Confirmation emails ("your application was sent…") are recognised too, and those jobs are kept out of your list.

## Configuration

Everything lives in `config.json`. It is re-read on every pass, so edits take effect without a restart. Start from [`config.example.json`](config.example.json), which describes a junior-to-mid backend engineer looking in India and Singapore.

### What you're looking for

| Key | Meaning |
|---|---|
| `title_include` | A job title must contain at least one of these (whole words, case-insensitive) |
| `title_exclude` | …and none of these. Use it for seniority (`senior`, `lead`) and fields you don't want |
| `skills` | `{"skill": points}`. Each skill found in the title or description adds its points to the Fit score. `go` / `golang` is matched carefully so the English word "go" doesn't count |
| `min_score` | Matches below this Fit score are dropped. Every job starts at 20, or 45 if it came from a job alert you set up yourself, and gets +15 when it asks for ≤1 year of experience or +8 for more |
| `max_required_years` | Jobs asking for more years than this go to *Stretch roles* instead of the main list… |
| `stretch_years` | …and jobs asking for more than this are dropped |
| `max_posting_age_days` | Ignore postings older than this |
| `min_salary_lpa` | Drop jobs whose stated **INR** salary tops out below this many lakhs per annum (jobs with no stated salary, or another currency, are kept) |
| `priority_salary_lpa` | Rank jobs paying at least this (INR, LPA) higher on *Apply today*. Leave it out to turn this off |
| `target_employers` | Companies you most want. They get a ★ and rank higher on *Apply today* |
| `notify_min_fit` | macOS notification for new matches at or above this Fit |
| `poll_minutes` | How often the background run checks |

### Regions

Each entry in `regions` is a separate search with its own workbook:

```json
{
  "key": "sg",
  "name": "Singapore",
  "excel_file": "Jobs_Singapore.xlsx",
  "locations": ["singapore"],
  "boards": [["LinkedIn - Backend, Singapore", "https://www.linkedin.com/jobs/search/?keywords=backend&location=Singapore"]]
}
```

- `locations`: a job's location must contain one of these words. Add `"remote"` to include remote roles.
- `boards`: links listed on the *Job boards* sheet. These are reminders for you; the bot doesn't read them.
- `key` must stay the same once you've started, because found jobs are stored under it.

### Work permits (optional)

If you're looking abroad, add `work_permit` to a region. The workbook then gets a **Work-permit notes** column that flags postings mentioning visa sponsorship, relocation, a required right to work, or a required language, and it drops postings that say they're for citizens only.

```json
"work_permit": {
  "country": "Indonesia",
  "nationality": "Indonesian",
  "language": "Indonesian",
  "citizens_only_terms": ["WNI", "warga negara indonesia"],
  "permit_terms": ["KITAS", "WNA"],
  "language_terms": ["bahasa indonesia"],
  "local_words": ["dan", "yang", "dengan", "untuk", "atau", "pengalaman", "minimal", "memiliki", "mampu", "dalam"]
}
```

All keys are optional:

- `country` flags postings asking for an existing "right to work in …".
- `nationality` drops "… citizens only" postings.
- `language` flags postings asking for fluency.
- The `*_terms` lists add local phrases to each check.
- `local_words` are common words of the local language. When 8 or more appear, the posting is flagged as written in that language.

### Job-alert emails

| Key | Meaning |
|---|---|
| `alert_senders` | `[label, sender address or domain, default location]`. The default location fills in when a card doesn't give one, e.g. `"India"` for Naukri |
| `alert_days` | How far back to look on the first read |

## Connecting Gmail (optional)

The bot reads alert emails over IMAP. It works with **Gmail only**, because it uses Gmail's search syntax.

1. Turn on [2-Step Verification](https://myaccount.google.com/security) for the Google account.
2. Create an [app password](https://myaccount.google.com/apppasswords). It's 16 characters.
3. Run `python3 jobbot.py setup-email` and enter the address.
   - **macOS:** paste the app password when asked. It goes straight into Keychain, never into a file.
   - **Linux / Windows:** set the `JOBBOT_EMAIL_PASSWORD` environment variable to the app password instead. It also works on macOS if you'd rather not use Keychain.
4. `python3 jobbot.py email-check` shows what it can read from your alerts, without saving anything.

Alerts take a day or two to start arriving after you create them.

## Files it creates

All of these are in `.gitignore`, so they stay out of git:

| File | What |
|---|---|
| `config.json` | Your settings |
| `*.xlsx` | Your workbooks |
| `.jobbot_state.json` | Every match found so far, with your statuses and notes |
| `jobbot.log` | Log of the background run |
| `.jobbot.pid` | Process id of the background run |
| `.jobbot_email` | The Gmail address only; never the password |

To start over for a region, stop the bot and delete its workbook and `.jobbot_state.json`.

## Running it at login (macOS)

`python3 jobbot.py start` keeps running until you stop it or log out. While the Mac sleeps it skips passes, logs that, and catches up when it wakes. To start it automatically at login, add a Login Item or a `launchd` agent that runs `python3 /path/to/jobbot.py run`.

## Notes

- Respect each site's terms. The bot only calls public listing endpoints, at a gentle pace (one pass an hour by default).
- Feeds break sometimes: a company moves to another ATS, or a board changes its page. A failing feed is logged in `jobbot.log` and skipped. The others keep working.
- Scoring is plain keyword matching, so tune `skills`, `title_include` and `title_exclude` to your own profile.

## License

[MIT](LICENSE)
