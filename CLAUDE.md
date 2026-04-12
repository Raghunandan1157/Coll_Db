# CLAUDE.md — Coll_Db EC2 Dashboard

## Project
NLPL Employee Dashboard hosted on AWS EC2 at `growwithme.navachetanalivelihoods.com`.
- **Frontend**: Static HTML/CSS/JS served by Apache from `/var/www/html/coll-db/` (symlinked to `~/Coll_Db/`)
- **Backend**: Node.js + Express + PostgreSQL on port 3000, proxied by Apache
- **Process Manager**: PM2 (`pm2 restart all` to reload server)

## EC2 Access
```
SSH_KEY=~/.ssh/aws-ec2.pem
EC2_HOST=ec2-user@52.66.163.52
```

## Deploy Rules
- Do NOT auto-deploy after edits. Only edit files locally.
- Deploy ONLY when the user explicitly says to **commit and push to GitHub**.
- When the user says to commit/push, do ALL three steps in sequence:
  1. `git add` + `git commit` + `git push origin master`
  2. Rsync to EC2:
     ```bash
     rsync -avz --exclude='node_modules' --exclude='.git' --exclude='data' --exclude='*.bak*' \
       -e "ssh -i ~/.ssh/aws-ec2.pem" \
       ./ ec2-user@52.66.163.52:~/Coll_Db/
     ```
  3. If `server/index.js` was changed, restart PM2:
     ```bash
     ssh -i ~/.ssh/aws-ec2.pem ec2-user@52.66.163.52 "pm2 restart all"
     ```
- For HTML/CSS/JS-only changes, rsync is enough — no PM2 restart needed.
- For server/index.js changes, MUST restart PM2 after rsync.

## Key Files
| File | Purpose |
|------|---------|
| `employee.html` | Main dashboard (Collection, Portfolio, Disbursement, Daily Reports tabs) |
| `daily-reports.html` | Daily Plan standalone app (also embedded in employee.html) |
| `contacts.html` | Employee directory |
| `locator.html` | Branch locator map |
| `server/index.js` | Node.js API server (PostgreSQL queries, upload endpoints) |
| `js/collection.js` | Collection tab logic + daily date routing |
| `js/employee.js` | Tab switching, role detection, date handling |
| `js/portfolio.js` | Portfolio tab (month-wise data) |
| `js/disbursement.js` | Disbursement tab |
| `js/daily-reports.js` | Daily Plan app JS (Neon replaced with EC2 API) |
| `css/desktop.css` | Desktop sidebar + responsive layout |

## Database (EC2 PostgreSQL)
- **Host**: 127.0.0.1:5432
- **User**: Raghunandan1157
- **Password**: raghu
- **Database**: postgres

### Tables
- `employee_performance` — Main dashboard data (latest upload via /api/upload)
- `daily_performance` — Per-date daily data (uploaded via /api/upload-daily)
- `daily_reports` — Daily Plan targets entered by BMs
- `daily_reports_achievements` — Daily Plan achievements entered by BMs
- `employees`, `branches`, `districts`, `regions`, `product_types` — Hierarchy tables

## Collection Data Processing
The EOD processor (`eod_processor.py` in the local Flask project) uses two formulas:
- **Regular dates**: `Sum(No of Regular Demand)` where FTOD & DPD Group excludes "1-30"
- **Month-end (last day of month)**: Same but ALSO excludes "31-60"

Auto-detected via `calendar.monthrange()`. See `batch_upload.py` for automated processing.

## Cache Busting
When editing JS/CSS, bump the version in the HTML file:
```
js/collection.js?v=33  →  js/collection.js?v=34
```

## Roles
- **CEO**: Sees all data, all tabs, report builder
- **RM**: Regional Manager — sees their region
- **DM**: Division Manager — sees their division
- **BM**: Branch Manager — sees their branch, can enter daily plans
- **FO**: Field Officer — sees their own data
