#!/usr/bin/env bash
set -euo pipefail

if [[ "${EUID}" -ne 0 ]]; then
  echo "Run this script with sudo."
  exit 1
fi

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_USER="${SUDO_USER:-$(stat -c '%U' "${PROJECT_DIR}")}"
APP_GROUP="$(id -gn "${APP_USER}")"
ENV_DIR="/etc/autoaudiobook"
ENV_FILE="${ENV_DIR}/autoaudiobook.env"
SERVICE_NAME="autoaudiobook"
SERVICE_PATH="/etc/systemd/system/${SERVICE_NAME}.service"
VENV_DIR="${PROJECT_DIR}/.venv"
PYTHON_BIN="${VENV_DIR}/bin/python"
UVICORN_BIN="${PROJECT_DIR}/.venv/bin/uvicorn"
DEFAULT_GEMINI_TEXT_MODEL="gemini-2.5-flash"
DEFAULT_GEMINI_TTS_MODEL="gemini-3.1-flash-tts-preview"

if ! id "${APP_USER}" >/dev/null 2>&1; then
  echo "App user ${APP_USER} does not exist."
  exit 1
fi

escape_env_value() {
  printf '%s' "$1" | sed -e 's/\\/\\\\/g' -e 's/"/\\"/g'
}

existing_api_key=""
if [[ -f "${ENV_FILE}" ]]; then
  existing_api_key="$(sed -n 's/^GEMINI_API_KEY="\(.*\)"$/\1/p' "${ENV_FILE}" | sed 's/\\"/"/g; s/\\\\/\\/g' | head -n 1)"
fi

echo "Installing system packages..."
apt-get update
apt-get install -y \
  ffmpeg \
  git \
  python3-pip \
  python3-venv

if [[ ! -d "${VENV_DIR}" ]]; then
  echo "Creating virtual environment..."
  runuser -u "${APP_USER}" -- python3 -m venv "${VENV_DIR}"
fi

echo "Installing Python dependencies..."
runuser -u "${APP_USER}" -- "${PYTHON_BIN}" -m pip install --upgrade pip
runuser -u "${APP_USER}" -- "${PYTHON_BIN}" -m pip install -r "${PROJECT_DIR}/requirements.txt"

echo
if [[ -n "${existing_api_key}" ]]; then
  echo "Press Enter to keep the existing Gemini API key."
fi

while true; do
  read -rsp "Gemini API key: " gemini_api_key
  echo
  if [[ -n "${gemini_api_key}" ]]; then
    break
  fi
  if [[ -n "${existing_api_key}" ]]; then
    gemini_api_key="${existing_api_key}"
    break
  fi
  echo "Gemini API key is required."
done

install -d -m 750 -o root -g "${APP_GROUP}" "${ENV_DIR}"

tmp_env_file="$(mktemp)"
cat > "${tmp_env_file}" <<EOF
GEMINI_API_KEY="$(escape_env_value "${gemini_api_key}")"
GEMINI_TEXT_MODEL="${DEFAULT_GEMINI_TEXT_MODEL}"
GEMINI_TTS_MODEL="${DEFAULT_GEMINI_TTS_MODEL}"
EOF
chown root:"${APP_GROUP}" "${tmp_env_file}"
chmod 640 "${tmp_env_file}"
mv "${tmp_env_file}" "${ENV_FILE}"

if [[ ! -x "${UVICORN_BIN}" ]]; then
  echo "Missing ${UVICORN_BIN}. Dependency install did not complete successfully."
  exit 1
fi

cat > "${SERVICE_PATH}" <<EOF
[Unit]
Description=AutoAudioBook Uvicorn Service
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=${APP_USER}
WorkingDirectory=${PROJECT_DIR}
EnvironmentFile=-${ENV_FILE}
ExecStart=${UVICORN_BIN} app:app --host 0.0.0.0 --port 8000
Restart=always
RestartSec=5
Environment=PYTHONUNBUFFERED=1

[Install]
WantedBy=multi-user.target
EOF

echo "Installing systemd service..."
systemctl daemon-reload
systemctl enable --now "${SERVICE_NAME}"
systemctl status "${SERVICE_NAME}" --no-pager

echo
echo "Install complete."
echo "Protected environment file: ${ENV_FILE}"
echo "Service status: sudo systemctl status ${SERVICE_NAME}"
