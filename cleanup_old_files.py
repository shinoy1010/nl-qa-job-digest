#!/usr/bin/env python3
"""Removes generated CV/cover-letter files whose last git commit is older
than MAX_AGE_DAYS. Age is measured via `git log` on the file's path, not
filesystem mtime -- mtime resets on every fresh checkout (e.g. in CI) and
can't be trusted to reflect when a file was actually generated.

A file's "last commit" also naturally refreshes if the same job reappears
in a later scrape and the packet gets regenerated -- so this purges packets
for jobs that haven't resurfaced in 30+ days, not just old calendar dates.
"""
import os
import subprocess
import time

MAX_AGE_DAYS = 30
DIRS = ["Tailored_CVs", "Cover_Letters"]


def last_commit_epoch(path):
    result = subprocess.run(
        ["git", "log", "-1", "--format=%ct", "--", path],
        capture_output=True, text=True, check=True,
    )
    out = result.stdout.strip()
    return int(out) if out else None


def main():
    cutoff = time.time() - MAX_AGE_DAYS * 86400
    removed = []

    for d in DIRS:
        if not os.path.isdir(d):
            continue
        for fname in sorted(os.listdir(d)):
            path = os.path.join(d, fname)
            if not os.path.isfile(path):
                continue
            ts = last_commit_epoch(path)
            if ts is None:
                continue  # not yet committed -- leave it for this run
            if ts < cutoff:
                subprocess.run(["git", "rm", "-q", path], check=True)
                removed.append(path)

    if removed:
        print(f"🗑️  Removed {len(removed)} file(s) older than {MAX_AGE_DAYS} days:")
        for p in removed:
            print(f"   - {p}")
    else:
        print(f"No CV/cover-letter files older than {MAX_AGE_DAYS} days.")

    return removed


if __name__ == "__main__":
    main()
