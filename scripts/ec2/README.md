# EC2 ops scripts (SSL watchdog)

Backup copies of files installed on EC2 (52.66.163.52). If the instance is
rebuilt, reinstall with the steps below.

## Files

| Repo file | EC2 location |
|---|---|
| `cert-watchdog.sh` | `/usr/local/bin/cert-watchdog.sh` (mode 755, run as root) |
| `cert-watchdog.systemd` | split into `/etc/systemd/system/cert-watchdog.service` + `.timer` |

## What it does

Twice daily (02:42 / 14:42 UTC): reads cert expiry, writes
`/home/ec2-user/cert-status.json` (served by `/api/cert-status`, shown as a
red banner to CEO in employee.html when <14 days). If <14 days it also
force-runs `certbot renew` + reloads Apache (self-heal when the main
`certbot-renew.timer` silently fails).

## Reinstall

```bash
sudo cp cert-watchdog.sh /usr/local/bin/ && sudo chmod 755 /usr/local/bin/cert-watchdog.sh
# split cert-watchdog.systemd into the two unit files, then:
sudo systemctl daemon-reload
sudo systemctl enable --now cert-watchdog.timer
sudo systemctl enable --now certbot-renew.timer   # main auto-renew
```

## Related monitoring

- `certbot-renew.timer` — main Let's Encrypt auto-renew (systemd, daily)
- UptimeRobot monitor "GrowWithMe Dashboard" (account: navachetana.raghu@gmail.com)
  — emails on site down, 5-min checks
