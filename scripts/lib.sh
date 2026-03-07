#!/usr/bin/env bash
# Shared shell functions for install.sh and upgrade.sh.
# Sourced, not executed directly.

set -euo pipefail

LIB_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$LIB_DIR/.." && pwd)"

# ---------------------------------------------------------------------------
# Colors and output helpers
# ---------------------------------------------------------------------------

if [[ -t 1 ]]; then
  RED='\033[0;31m'
  GREEN='\033[0;32m'
  YELLOW='\033[0;33m'
  BLUE='\033[0;34m'
  BOLD='\033[1m'
  RESET='\033[0m'
else
  RED='' GREEN='' YELLOW='' BLUE='' BOLD='' RESET=''
fi

info()    { printf "${BLUE}i${RESET} %s\n" "$*"; }
warn()    { printf "${YELLOW}!${RESET} %s\n" "$*"; }
error()   { printf "${RED}x${RESET} %s\n" "$*" >&2; }
success() { printf "${GREEN}+${RESET} %s\n" "$*"; }
step()    { printf "\n${BOLD}> %s${RESET}\n" "$*"; }

# ---------------------------------------------------------------------------
# Flags
# ---------------------------------------------------------------------------

DRY_RUN=false

parse_flags() {
  for arg in "$@"; do
    case "$arg" in
      --dry-run) DRY_RUN=true ;;
    esac
  done
}

# ---------------------------------------------------------------------------
# open_url -- open a URL in the default browser, or print it
# ---------------------------------------------------------------------------

open_url() {
  local url="$1"
  if [[ "$DRY_RUN" == true ]]; then
    info "[dry-run] Would open: $url"
    return
  fi
  if command -v open &>/dev/null; then
    open "$url" 2>/dev/null || true
  elif command -v xdg-open &>/dev/null; then
    xdg-open "$url" 2>/dev/null || true
  else
    info "Open this URL in your browser:"
    printf "  %s\n" "$url"
  fi
}

# ---------------------------------------------------------------------------
# read_manifest -- parse setup-manifest.json via node
# ---------------------------------------------------------------------------

MANIFEST_SETUP_VERSION=0
MANIFEST_MIN_NODE_VERSION=18
MANIFEST_JSON=""

read_manifest() {
  local manifest="$REPO_ROOT/setup-manifest.json"
  if ! command -v node &>/dev/null; then
    error "Node.js is required but not found. Please install Node.js and try again."
    exit 1
  fi
  if [[ ! -f "$manifest" ]]; then
    error "setup-manifest.json not found at $manifest"
    exit 1
  fi

  MANIFEST_JSON="$(cat "$manifest")"

  eval "$(_node -e "
    const m = JSON.parse(process.argv[1]);
    console.log('MANIFEST_SETUP_VERSION=' + m.setup_version);
    console.log('MANIFEST_MIN_NODE_VERSION=' + m.min_node_version);
  " "$MANIFEST_JSON")"
}

# ---------------------------------------------------------------------------
# check_prerequisites -- validate node version and npm
# ---------------------------------------------------------------------------

check_prerequisites() {
  step "Checking prerequisites"

  local node_major
  node_major=$(_node -e 'console.log(process.versions.node.split(".")[0])')
  if (( node_major < MANIFEST_MIN_NODE_VERSION )); then
    error "Node.js v${node_major} found, but v${MANIFEST_MIN_NODE_VERSION}+ is required."
    exit 1
  fi
  success "Node.js $(node --version)"

  if ! command -v npm &>/dev/null; then
    error "npm is not installed. Please install npm and try again."
    exit 1
  fi
  success "npm v$(npm --version)"
}

# ---------------------------------------------------------------------------
# resolve_credential_paths -- honor env vars, derive sidecar path
# ---------------------------------------------------------------------------

OAUTH_KEYFILE=""
CREDENTIALS_FILE=""
CREDENTIALS_DIR=""
SIDECAR_FILE=""

resolve_credential_paths() {
  OAUTH_KEYFILE="${GDRIVE_OAUTH_PATH:-$REPO_ROOT/credentials/gcp-oauth.keys.json}"
  CREDENTIALS_FILE="${GDRIVE_CREDENTIALS_PATH:-$REPO_ROOT/credentials/.gdrive-server-credentials.json}"
  CREDENTIALS_DIR="$(dirname "$CREDENTIALS_FILE")"
  SIDECAR_FILE="$CREDENTIALS_DIR/.gdrive-setup.json"

  local oauth_dir
  oauth_dir="$(dirname "$OAUTH_KEYFILE")"

  if [[ "$DRY_RUN" == true ]]; then
    info "[dry-run] Credentials dir: $CREDENTIALS_DIR"
    if [[ "$oauth_dir" != "$CREDENTIALS_DIR" ]]; then
      info "[dry-run] OAuth keyfile dir: $oauth_dir"
    fi
  else
    mkdir -p "$CREDENTIALS_DIR"
    mkdir -p "$oauth_dir"
  fi
}

# ---------------------------------------------------------------------------
# read_sidecar / write_sidecar
# ---------------------------------------------------------------------------

SIDECAR_SETUP_VERSION=0
SIDECAR_AUTH_COMPLETED=false
SIDECAR_REAUTH_PENDING=false

read_sidecar() {
  if [[ -f "$SIDECAR_FILE" ]]; then
    eval "$(_node -e "
      const s = JSON.parse(require('fs').readFileSync(process.argv[1], 'utf-8'));
      console.log('SIDECAR_SETUP_VERSION=' + (s.setup_version || 0));
      console.log('SIDECAR_AUTH_COMPLETED=' + (s.auth_completed || false));
      console.log('SIDECAR_REAUTH_PENDING=' + (s.reauth_pending || false));
    " "$SIDECAR_FILE")"
  else
    SIDECAR_SETUP_VERSION=0
    SIDECAR_AUTH_COMPLETED=false
    SIDECAR_REAUTH_PENDING=false
  fi
}

write_sidecar() {
  local version="${1:-$MANIFEST_SETUP_VERSION}"
  local auth_completed="${2:-true}"
  local reauth_pending="${3:-false}"

  if [[ "$DRY_RUN" == true ]]; then
    info "[dry-run] Would write sidecar: version=$version auth_completed=$auth_completed reauth_pending=$reauth_pending"
    return
  fi

  _node -e "
    const fs = require('fs');
    const data = {
      setup_version: parseInt(process.argv[1], 10),
      auth_completed: process.argv[2] === 'true',
      reauth_pending: process.argv[3] === 'true',
      updated_at: new Date().toISOString()
    };
    fs.writeFileSync(process.argv[4], JSON.stringify(data, null, 2) + '\n');
  " "$version" "$auth_completed" "$reauth_pending" "$SIDECAR_FILE"
}

# ---------------------------------------------------------------------------
# pending_migrations -- returns count and prints migration JSON lines
# ---------------------------------------------------------------------------

pending_migration_count() {
  _node -e "
    const m = JSON.parse(process.argv[1]);
    const v = parseInt(process.argv[2], 10);
    console.log(m.migrations.filter(mig => v < mig.from_version_below).length);
  " "$MANIFEST_JSON" "$SIDECAR_SETUP_VERSION"
}

# Prints one JSON object per line for each pending migration
pending_migration_records() {
  _node -e "
    const m = JSON.parse(process.argv[1]);
    const v = parseInt(process.argv[2], 10);
    m.migrations
      .filter(mig => v < mig.from_version_below)
      .forEach(mig => console.log(JSON.stringify(mig)));
  " "$MANIFEST_JSON" "$SIDECAR_SETUP_VERSION"
}

# ---------------------------------------------------------------------------
# prompt_yes_no -- interactive y/n helper; returns 0 for yes, 1 for no
# ---------------------------------------------------------------------------

prompt_yes_no() {
  local prompt="$1"
  local default="${2:-y}"

  if [[ "$DRY_RUN" == true ]]; then
    info "[dry-run] Would prompt: $prompt (default: $default)"
    return 0
  fi

  local yn
  if [[ "$default" == "y" ]]; then
    printf "%s [Y/n] " "$prompt"
  else
    printf "%s [y/N] " "$prompt"
  fi

  read -r yn
  yn="${yn:-$default}"
  [[ "$yn" =~ ^[Yy] ]]
}

# ---------------------------------------------------------------------------
# run_or_dry -- execute or print what would run
# ---------------------------------------------------------------------------

run_or_dry() {
  if [[ "$DRY_RUN" == true ]]; then
    info "[dry-run] Would run: $*"
  else
    "$@"
  fi
}

# ---------------------------------------------------------------------------
# GCP project helpers
# ---------------------------------------------------------------------------

GCP_PROJECT_ID=""

ask_gcp_project() {
  if [[ -n "$GCP_PROJECT_ID" ]]; then
    return
  fi
  if [[ "$DRY_RUN" == true ]]; then
    info "[dry-run] Would prompt for GCP project ID."
    return
  fi

  printf "\n"
  info "Enter your Google Cloud project ID (or paste a console URL containing it)."
  info "Find it at https://console.cloud.google.com/ -- shown in the project selector at the top."
  printf "  Project ID: "
  read -r input

  # Extract project ID from a URL if pasted
  if [[ "$input" == *"project="* ]]; then
    GCP_PROJECT_ID=$(printf '%s' "$input" | sed 's/.*project=\([^&]*\).*/\1/')
  else
    GCP_PROJECT_ID="$input"
  fi

  if [[ -z "$GCP_PROJECT_ID" ]]; then
    warn "No project ID provided. Cloud Console links will not be project-scoped."
  else
    success "Using project: $GCP_PROJECT_ID"
  fi
}

build_console_url() {
  local link_key="$1"
  local api_id="${2:-}"

  local url
  url=$(_node -e "
    const m = JSON.parse(process.argv[1]);
    let url = m.console_links[process.argv[2]] || '';
    const apiId = process.argv[3];
    if (apiId) url = url.replace('{api_id}', apiId);
    console.log(url);
  " "$MANIFEST_JSON" "$link_key" "$api_id")

  if [[ -n "$GCP_PROJECT_ID" ]]; then
    if [[ "$url" == *"?"* ]]; then
      url="${url}&project=${GCP_PROJECT_ID}"
    else
      url="${url}?project=${GCP_PROJECT_ID}"
    fi
  fi

  printf '%s' "$url"
}

# ---------------------------------------------------------------------------
# _node -- run node with colors disabled for clean machine-readable output
# ---------------------------------------------------------------------------

_node() {
  FORCE_COLOR=0 node "$@"
}

# ---------------------------------------------------------------------------
# wait_for_enter
# ---------------------------------------------------------------------------

wait_for_enter() {
  if [[ "$DRY_RUN" == true ]]; then
    return
  fi
  local msg="${1:-Press Enter to continue...}"
  printf "\n  ${BOLD}%s${RESET} " "$msg"
  read -r
}

# ---------------------------------------------------------------------------
# print_mcp_config -- print ready-to-copy MCP client configurations
# ---------------------------------------------------------------------------

print_mcp_config() {
  local abs_index="$REPO_ROOT/dist/index.js"

  step "Add to your MCP client"

  printf "\n  ${BOLD}Claude Code CLI:${RESET}\n"
  printf '  claude mcp add --scope user wagnerlabs-gdrive -- node "%s"\n' "$abs_index"

  printf "\n  ${BOLD}Cursor (.cursor/mcp.json):${RESET}\n"
  printf '  {\n'
  printf '    "mcpServers": {\n'
  printf '      "gdrive": {\n'
  printf '        "command": "node",\n'
  printf '        "args": ["%s"]\n' "$abs_index"
  printf '      }\n'
  printf '    }\n'
  printf '  }\n'

  printf "\n  ${BOLD}Claude Desktop (claude_desktop_config.json):${RESET}\n"
  printf '  {\n'
  printf '    "mcpServers": {\n'
  printf '      "gdrive": {\n'
  printf '        "command": "node",\n'
  printf '        "args": ["%s"]\n' "$abs_index"
  printf '      }\n'
  printf '    }\n'
  printf '  }\n'
}
