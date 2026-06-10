#!/bin/bash
# cert-watchdog: daily SSL expiry check + self-heal + status JSON for dashboard banner
CERT=/etc/letsencrypt/live/growwithme.navachetanalivelihoods.com/fullchain.pem
STATUS=/home/ec2-user/cert-status.json

days_left() {
  local end epoch
  end=$(openssl x509 -enddate -noout -in "$CERT" 2>/dev/null | cut -d= -f2) || return 1
  epoch=$(date -d "$end" +%s) || return 1
  echo $(( (epoch - $(date +%s)) / 86400 ))
}

d=$(days_left)
[ -z "$d" ] && d=-1

# Self-heal: renewal timer should have handled this; force attempt if close
if [ "$d" -lt 14 ]; then
  certbot renew --quiet && systemctl reload httpd
  d=$(days_left)
  [ -z "$d" ] && d=-1
fi

printf "{\"days_left\": %s, \"checked_at\": \"%s\"}\n" "$d" "$(date -u +%FT%TZ)" > "$STATUS"
chown ec2-user:ec2-user "$STATUS"
