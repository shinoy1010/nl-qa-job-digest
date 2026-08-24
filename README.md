# Job Automation Engine

Fully self-contained NL job-search pipeline. Runs entirely outside Claude:
scrapes job listings via **Apify**, scores and tailors a resume + cover
letter per listing using **Google Gemini**, and generates a styled local
HTML digest plus per-job `.docx` files. No dependency on Claude, claude.ai,
or any Claude Code subscription.

## Setup

```bash
pip install apify-client google-genai python-docx pandas
cp .env.example .env
# then edit .env and fill in your own keys
```

`master_resume.txt` holds the master resume text Gemini tailors against --
edit it directly if your background changes.

### Email delivery (optional)

To have each successful run email the digest to you, set `EMAIL_SENDER`,
`EMAIL_APP_PASSWORD`, and `EMAIL_RECIPIENT` in `.env` (or as GitHub Actions
secrets for the scheduled workflow). `EMAIL_SENDER` must be a Gmail address
with an **App Password** -- not your regular Gmail password:

1. Enable 2-Step Verification on the sender Gmail account (Google Account ->
   Security).
2. Go to [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords),
   create an app password (name it e.g. "job-automation-engine"), and copy
   the 16-character code.
3. Use that code as `EMAIL_APP_PASSWORD`.

If these aren't set, the pipeline still runs normally and just skips the
email step.

## Run

```bash
python3 app_engine.py
```

Each run:
1. Triggers the configured Apify scraping task and pulls the job listings.
2. Filters out listings that require Dutch (unless it's flagged as a plus).
3. Sends each remaining listing + your master resume to Gemini, which scores
   the fit (0-100) and drafts a tailored summary, bullets, and cover letter.
4. Skips anything scoring below 40.
5. Generates a tailored one-page CV (`Tailored_CVs/`) and cover letter
   (`Cover_Letters/`) per qualifying job.
6. Writes `QA_roles_<date>.csv` (raw data) and `reports/QA_Job_Digest_<date>.html`
   (the readable report -- open it in a browser; its Resume/Cover letter
   links point straight at the generated files on disk).
7. Emails the HTML digest to `EMAIL_RECIPIENT` if email is configured (see
   below).

Nothing is submitted anywhere automatically -- every document is a draft for
manual review.

## Notes

- API keys live in `.env` (gitignored) -- never commit real keys. `.env.example`
  documents the expected variables.
- Generated output (`Tailored_CVs/`, `Cover_Letters/`, `reports/`, CSVs) is
  gitignored by default since it's regenerated on every run and is mostly
  binary `.docx` files that don't diff well. Remove those lines from
  `.gitignore` if you'd rather keep application history versioned in git.
