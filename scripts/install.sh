#!/usr/bin/env bash
# First-time setup walkthrough for gdrive-mcp.
# Usage: bash scripts/install.sh [--dry-run]

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
source "$SCRIPT_DIR/lib.sh"

parse_flags "$@"

main() {
  printf "\n${BOLD}Google Drive MCP Server -- Setup${RESET}\n"
  printf "================================\n"

  # -----------------------------------------------------------------------
  # 1. Read manifest and check prerequisites
  # -----------------------------------------------------------------------
  read_manifest
  check_prerequisites

  # -----------------------------------------------------------------------
  # 2. npm install && npm run build
  # -----------------------------------------------------------------------
  step "Installing dependencies and building"
  run_or_dry npm install --prefix "$REPO_ROOT"
  run_or_dry npm run --prefix "$REPO_ROOT" build

  # -----------------------------------------------------------------------
  # 3. Resolve credential paths
  # -----------------------------------------------------------------------
  resolve_credential_paths

  # -----------------------------------------------------------------------
  # 4. Google Cloud Console walkthrough (skip if OAuth keyfile exists)
  # -----------------------------------------------------------------------
  if [[ -f "$OAUTH_KEYFILE" ]]; then
    success "OAuth keyfile found at $OAUTH_KEYFILE -- skipping Cloud Console walkthrough."
  else
    cloud_console_walkthrough
  fi

  # -----------------------------------------------------------------------
  # 5. Run auth
  # -----------------------------------------------------------------------
  step "Authenticating with Google"
  if [[ "$DRY_RUN" == true ]]; then
    info "[dry-run] Would run: node \"$REPO_ROOT/dist/index.js\" auth"
  else
    if node "$REPO_ROOT/dist/index.js" auth; then
      success "Authentication complete!"
    else
      error "Authentication failed."
      info "You can retry later with: node \"$REPO_ROOT/dist/index.js\" auth"
      exit 1
    fi
  fi

  # -----------------------------------------------------------------------
  # 6. Write sidecar
  # -----------------------------------------------------------------------
  write_sidecar "$MANIFEST_SETUP_VERSION" "true" "false"

  # -----------------------------------------------------------------------
  # 7. Print MCP client config
  # -----------------------------------------------------------------------
  print_mcp_config

  # -----------------------------------------------------------------------
  # Done
  # -----------------------------------------------------------------------
  printf "\n"
  success "Setup complete! Add the MCP server config above to your preferred client."
}

cloud_console_walkthrough() {
  step "Google Cloud Console setup"
  info "We'll walk you through creating a Google Cloud project and OAuth credentials."
  info "Each step opens a browser link. Follow the instructions, then come back here."

  # a. Ask for GCP project
  ask_gcp_project

  # b. Create or select project
  step "Step 1: Create or select a Google Cloud project"
  info "If you already have a project, you can skip this step."
  local url
  url=$(build_console_url "project_create")
  info "Link: $url"
  open_url "$url"
  wait_for_enter

  # c. Enable required APIs
  step "Step 2: Enable required APIs"
  info "Enable each of the following APIs in your project:"
  printf "\n"
  while IFS=$'\t' read -r api_id api_name; do
    local api_url
    api_url=$(build_console_url "api_library" "$api_id")
    info "  $api_name"
    info "  $api_url"
    open_url "$api_url"
    printf "\n"
  done < <(_node -e "
    const m = JSON.parse(process.argv[1]);
    m.required_apis.forEach(a => console.log(a.id + '\t' + a.name));
  " "$MANIFEST_JSON")
  wait_for_enter "Press Enter after enabling all APIs..."

  # d. Configure OAuth consent screen
  step "Step 3: Configure the OAuth consent screen"
  info "1. Click 'Get started'"
  info "2. App name: gdrive-mcp"
  info "3. User support email: select your email"
  info "4. Audience: Internal (Workspace) or External (personal Gmail)"
  info "5. Contact email: your email"
  info "6. Finish and create"
  url=$(build_console_url "oauth_consent")
  info "Link: $url"
  open_url "$url"
  wait_for_enter

  # e. Add required scopes
  step "Step 4: Add required scopes"
  info "In the left sidebar, go to 'Data Access', then click 'Add or remove scopes'."
  info "Add these scopes:"
  printf "\n"
  while read -r scope; do
    info "  $scope"
  done < <(_node -e "
    const m = JSON.parse(process.argv[1]);
    m.required_scopes.forEach(s => console.log(s));
  " "$MANIFEST_JSON")
  printf "\n"
  info "Save when done."
  wait_for_enter

  # f. Create OAuth Desktop client
  step "Step 5: Create an OAuth Desktop client"
  info "1. Application type: Desktop app"
  info "2. Name: gdrive-mcp"
  info "3. Click 'Create'"
  url=$(build_console_url "oauth_client_create")
  info "Link: $url"
  open_url "$url"
  wait_for_enter

  # g. Download and place keyfile
  step "Step 6: Download the OAuth client JSON"
  info "Download the JSON file and save it to:"
  info "  $OAUTH_KEYFILE"
  printf "\n"

  while true; do
    wait_for_enter "Press Enter after saving the file..."
    if [[ "$DRY_RUN" == true ]]; then
      info "[dry-run] Skipping file existence check."
      break
    fi
    if [[ -f "$OAUTH_KEYFILE" ]]; then
      success "OAuth keyfile found!"
      break
    fi
    warn "File not found at: $OAUTH_KEYFILE"
    info "Please download the JSON and save it to that path."
  done
}

main
