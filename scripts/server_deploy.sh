#!/usr/bin/env bash
# Server-side deploy: install deps, run migrations, collect static, restart gunicorn.
# Invoked by .github/workflows/main.yml after `git reset --hard origin/main`.
set -euo pipefail

APP_DIR="/opt/gazebo_hr"
VENV="$APP_DIR/venv"

cd "$APP_DIR"

if [ ! -d "$VENV" ]; then
    # Django 6 requires Python >=3.12
    if command -v python3.12 >/dev/null 2>&1; then
        python3.12 -m venv "$VENV"
    else
        python3 -m venv "$VENV"
    fi
fi

"$VENV/bin/pip" install --quiet --upgrade pip wheel
"$VENV/bin/pip" install --quiet -r requirements.txt

if [ -f "$APP_DIR/.env" ]; then
    set -a
    # shellcheck disable=SC1091
    . "$APP_DIR/.env"
    set +a
fi

"$VENV/bin/python" manage.py migrate --noinput
"$VENV/bin/python" manage.py collectstatic --noinput

sudo systemctl restart gazebo_hr.service
sudo systemctl --no-pager status gazebo_hr.service | head -n 5
