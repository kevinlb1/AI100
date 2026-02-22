\# Windows setup (simple `main` branch workflow)



This guide is intentionally simple:

\- \*\*one branch only\*\* (`main`),

\- no PR-branch checkout,

\- and explicit recovery commands if files were deleted locally.



\## 0) Open Command Prompt in your repo folder



\## One-click option



You can also run `UPDATE\_MAIN\_AND\_RUN.bat` from this repo root. It executes the same sync/install/run steps automatically.



```bat

cd D:\\Dropbox\\Admin\\Programming\\AI100

```



If this is a fresh machine:



```bat

git clone https://github.com/kevinlb1/AI100.git

cd AI100

```



\## 1) Sync local files to exactly `origin/main`



```bat

git fetch origin --prune

git checkout main

git reset --hard origin/main

```



This is the key command set for everyday use. It restores tracked files and removes local drift.



\## 2) Verify the key files are present



```bat

dir

git --no-pager log --oneline -n 5

```



You should see at least:

\- `app.py`

\- `match.py`

\- `group\_assignments.tex`



\## 3) Create and activate virtual environment (Windows CMD)



```bat

python -m venv .venv

.venv\\Scripts\\activate

```



> `source .venv/bin/activate` is Linux/macOS syntax, not Windows CMD.



\## 4) Install dependencies



```bat

python -m pip install --upgrade pip

pip install ortools numpy matplotlib

```



\## 5) Run the app



```bat

python app.py

```



Open:

\- http://localhost:8000



\## Why this can still say "up to date"



`git pull` only downloads commits that already exist on your GitHub remote.

If agent commits were not actually pushed to a remote branch, your local repo cannot fetch them.

So yes: if commits are not on a remote branch, there is nothing for `pull` to retrieve.



\## How to actually get the latest agent changes



If `git pull` says "Already up to date" but `app.py` is still old, run:



```bat

GET\_AGENT\_CHANGES.bat

```



This script will:

1\. Sync your local `main` with `origin/main`.

2\. Check known remote branches for `app.py` and for the v2 marker (`v2-multiuser`).

3\. If v2 exists remotely, merge that branch into `main` and push.

4\. If v2 does \*\*not\*\* exist remotely, print diagnostics that prove no published v2 commit is available to merge yet.



\## Emergency local fix (no remote merge needed)



If you still see the old UI and just want this repo to use the v2 app immediately, run:



```bat

APPLY\_V2\_LOCALLY.bat

```



This script does not require merging a PR first. It searches local/remote branches for `app.py` containing `v2-multiuser`, restores `app.py` + related files into your current `main`, and creates a local commit if needed.



\## Troubleshooting



\- \*\*Do I need to manually accept every PR?\*\*

&nbsp; - Not if commits are already pushed to `main`.

&nbsp; - But if only a PR \*description\* exists and no remote branch has the commit, you must either:

&nbsp;   1. have the agent publish/push the branch, or

&nbsp;   2. apply the patch/file changes manually.



\- \*\*`GET\_AGENT\_CHANGES.bat` says no remote branch contains `v2-multiuser`\*\*

&nbsp; - That is a remote state issue, not a local checkout issue.

&nbsp; - In this case, your remotes simply do not yet have the v2 commit.

&nbsp; - Share the script diagnostics with the agent so the correct branch can be pushed/published.



\- \*\*`git pull` says "Already up to date" but `app.py` is still the old UI\*\*

&nbsp; - This means your local `main` and `origin/main` agree, but that remote branch may not contain the new commit yet.

&nbsp; - Check your current history and file marker:

&nbsp;   ```bat

&nbsp;   git --no-pager log --oneline -n 5

&nbsp;   findstr /N /C:"v2-multiuser" app.py

&nbsp;   ```

&nbsp; - If marker is missing, fetch all remotes and inspect available branches:

&nbsp;   ```bat

&nbsp;   git fetch --all --prune

&nbsp;   git branch -r

&nbsp;   ```

&nbsp; - Then either merge the branch containing the new app into `main` on GitHub, or check it out locally to test.



\- \*\*`python: can't open file ... app.py`\*\*

&nbsp; - You are not in a synced working tree. Re-run exactly:

&nbsp;   ```bat

&nbsp;   git fetch origin --prune

&nbsp;   git checkout main

&nbsp;   git reset --hard origin/main

&nbsp;   dir

&nbsp;   ```



\- \*\*You deleted files manually and pull did not bring them back\*\*

&nbsp; - Use the same hard reset above. `git pull` updates commits, but does not always repair local working-tree deletions the way users expect.



\- \*\*`fatal: unable to write new index file`\*\*

&nbsp; - Usually a local file-lock/permissions/disk issue. Try:

&nbsp;   ```bat

&nbsp;   del .git\\index.lock

&nbsp;   attrib -R .git\\index

&nbsp;   git status

&nbsp;   ```

&nbsp; - Then re-run:

&nbsp;   ```bat

&nbsp;   git reset --hard origin/main

&nbsp;   ```

&nbsp; - If needed, close editors/terminals that may lock `.git\\index`, and verify free disk space.



\- \*\*`No module named 'ortools'`\*\*

&nbsp; - Re-activate venv and verify package location:

&nbsp;   ```bat

&nbsp;   .venv\\Scripts\\activate

&nbsp;   python -m pip show ortools

&nbsp;   ```



\- \*\*Git opens a big pager screen\*\*

&nbsp; - Use no-pager form for inspection commands:

&nbsp;   ```bat

&nbsp;   git --no-pager log --oneline -n 5

&nbsp;   ```



