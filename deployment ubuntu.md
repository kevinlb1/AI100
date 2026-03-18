# Ubuntu deployment

This repo can be run in two ways on Ubuntu:

- simple/local: `./run_ubuntu.sh`
- production web app at `/AI100`: `nginx + gunicorn + systemd`

If you want the app at `http://[machine]/AI100`, use the production setup. The app now supports a URL base path through `APP_BASE_PATH=/AI100`.

## Recommended production layout

- `nginx` listens on port `80`
- `nginx` proxies `/AI100/` to `gunicorn` on `127.0.0.1:8000`
- `gunicorn` should run with a single worker for this app's threaded solver jobs
- `systemd` keeps `gunicorn` running
- the app stores its SQLite database in `match_app.db` in the repo root

## 1. Clone the repo

```bash
git clone https://github.com/kevinlb1/AI100.git
cd AI100
```

## 2. Install Ubuntu packages and Python dependencies

```bash
chmod +x install_ubuntu.sh run_ubuntu.sh install_server_ubuntu.sh run_server.sh
./install_server_ubuntu.sh
```

That script:

- installs `python3`, `python3-venv`, `python3-pip`, and `nginx`
- creates `.venv`
- installs the app requirements
- installs `gunicorn`

## 3. Test the server locally

```bash
APP_BASE_PATH=/AI100 ./run_server.sh
```

Then from the Ubuntu machine itself:

```bash
curl http://127.0.0.1:8000/
```

Stop it with `Ctrl+C`.

## 4. Install the systemd service

Copy the example service and replace the placeholders:

- `__APP_USER__` with the Linux user that should run the app
- `__REPO_DIR__` with the absolute path to this repo on Ubuntu

Example if the repo lives at `/opt/AI100` and the app user is `kevin`:

```bash
sed \
  -e 's|__APP_USER__|kevin|g' \
  -e 's|__REPO_DIR__|/opt/AI100|g' \
  ai100.service.example | sudo tee /etc/systemd/system/ai100.service >/dev/null
```

Enable and start it:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now ai100
sudo systemctl status ai100
```

## 5. Install the nginx site

Copy the example config:

```bash
sudo cp nginx-ai100.conf.example /etc/nginx/sites-available/ai100
sudo ln -sf /etc/nginx/sites-available/ai100 /etc/nginx/sites-enabled/ai100
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl reload nginx
```

This config does two important things:

- redirects `/AI100` to `/AI100/`
- strips the `/AI100/` prefix before proxying to the Python app

## 6. Open the app

Use:

- `http://[machine]/AI100`

Examples:

- `http://192.168.1.50/AI100`
- `http://ubuntu-box/AI100`

## Logs and operations

Check the app logs:

```bash
sudo journalctl -u ai100 -f
```

Restart after code changes:

```bash
sudo systemctl restart ai100
```

Reload nginx after config changes:

```bash
sudo systemctl reload nginx
```

## Simple non-production mode

If you only want to run it directly without nginx:

```bash
./install_ubuntu.sh
./run_ubuntu.sh
```

That serves the app directly on:

- `http://127.0.0.1:8000`

## Notes

- `wsgiref` in `app.py` is fine for local testing, but `gunicorn + nginx` is the correct deployment model for a stable Ubuntu web app.
- The default `gunicorn.conf.py` now uses `1` worker because live solver runs are managed by background threads inside the web process. If you override `GUNICORN_WORKERS`, keep it at `1` unless the run orchestration is redesigned for multi-process execution.
- If Ubuntu firewall rules are enabled, allow HTTP:

```bash
sudo ufw allow 80/tcp
```

- If you want to expose the backend port directly for debugging, the `gunicorn` config uses `127.0.0.1:8000` by default, so it is not reachable externally unless you change it.
