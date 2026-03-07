#!/usr/bin/env bash
# Post-pull upgrade flow for gdrive-mcp.
# Usage: bash scripts/upgrade.sh [--dry-run]

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
source "$SCRIPT_DIR/lib.sh"

parse_flags "$@"

main() {
  printf "\n${BOLD}Google Drive MCP Server -- Upgrade${RESET}\n"
  printf "===================================\n"

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
  # 3. Resolve credential paths and read sidecar
  # -----------------------------------------------------------------------
  resolve_credential_paths
  read_sidecar

  # -----------------------------------------------------------------------
  # 4. Handle pending reauth from a previous upgrade
  # -----------------------------------------------------------------------
  if [[ "$SIDECAR_REAUTH_PENDING" == "true" ]]; then
    step "Pending re-authentication"
    warn "A previous upgrade required re-authentication that was deferred."
    handle_reauth
  fi

  # -----------------------------------------------------------------------
  # 5. Check for pending migrations
  # -----------------------------------------------------------------------
  local count
  count=$(pending_migration_count)

  if (( count == 0 )); then
    # Persist any reauth state change (e.g. user just completed deferred reauth)
    write_sidecar "$MANIFEST_SETUP_VERSION" "true" "$SIDECAR_REAUTH_PENDING"
    printf "\n"
    success "You're up to date (setup_version=$MANIFEST_SETUP_VERSION)."
    if [[ "$SIDECAR_REAUTH_PENDING" == "true" ]]; then
      warn "Re-authentication is still pending. Run: ./scripts/upgrade.sh"
    fi
    return
  fi

  if (( count > 0 )); then
    step "Applying $count migration(s)"
    local reauth_needed=false
    local migration_skipped=false

    while read -r migration_json; do
      local notice requires_reauth added_apis_count added_scopes_count
      notice=$(_node -e "console.log(JSON.parse(process.argv[1]).notice)" "$migration_json")
      requires_reauth=$(_node -e "console.log(JSON.parse(process.argv[1]).requires_reauth || false)" "$migration_json")
      added_apis_count=$(_node -e "const m=JSON.parse(process.argv[1]); console.log((m.added_apis||[]).length)" "$migration_json")
      added_scopes_count=$(_node -e "const m=JSON.parse(process.argv[1]); console.log((m.added_scopes||[]).length)" "$migration_json")

      # Show notice
      printf "\n"
      warn "$notice"

      # Handle added APIs
      if (( added_apis_count > 0 )); then
        ask_gcp_project
        printf "\n"
        info "New APIs to enable in your Google Cloud project:"
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
          const manifest = JSON.parse(process.argv[2]);
          (m.added_apis || []).forEach(id => {
            const api = manifest.required_apis.find(a => a.id === id);
            console.log(id + '\t' + (api ? api.name : id));
          });
        " "$migration_json" "$MANIFEST_JSON")

        if prompt_yes_no "Have you enabled the new APIs?"; then
          success "APIs confirmed."
        else
          migration_skipped=true
          warn "Skipped. Re-run './scripts/upgrade.sh' after enabling the APIs."
        fi
      fi

      # Handle added scopes
      if (( added_scopes_count > 0 )); then
        printf "\n"
        info "New OAuth scopes added:"
        while read -r scope; do
          info "  $scope"
        done < <(_node -e "
          const m = JSON.parse(process.argv[1]);
          (m.added_scopes || []).forEach(s => console.log(s));
        " "$migration_json")
      fi

      # Handle reauth
      if [[ "$requires_reauth" == "true" ]]; then
        reauth_needed=true
      fi
    done < <(pending_migration_records)

    # Process reauth after all migrations are shown
    if [[ "$reauth_needed" == true ]]; then
      printf "\n"
      warn "Re-authentication is required to grant the new scopes."
      handle_reauth
    fi

    # If the user skipped console steps, don't advance the version so the
    # next run re-prompts.
    if [[ "$migration_skipped" == true ]]; then
      write_sidecar "$SIDECAR_SETUP_VERSION" "true" "$SIDECAR_REAUTH_PENDING"
      printf "\n"
      warn "Some migration steps were skipped. Version not advanced -- re-run './scripts/upgrade.sh' when ready."
      return
    fi
  fi

  # -----------------------------------------------------------------------
  # 6. Write sidecar
  # -----------------------------------------------------------------------
  write_sidecar "$MANIFEST_SETUP_VERSION" "true" "$SIDECAR_REAUTH_PENDING"

  # -----------------------------------------------------------------------
  # 7. Summary
  # -----------------------------------------------------------------------
  printf "\n"
  success "Upgrade complete (setup_version=$MANIFEST_SETUP_VERSION)."
  if [[ "$SIDECAR_REAUTH_PENDING" == "true" ]]; then
    warn "Re-authentication was deferred. Run this to complete it later:"
    info "  rm \"$CREDENTIALS_FILE\" && node \"$REPO_ROOT/dist/index.js\" auth"
    info "  Then run: ./scripts/upgrade.sh"
  fi
}

handle_reauth() {
  if prompt_yes_no "Re-authenticate now?"; then
    if [[ "$DRY_RUN" == true ]]; then
      info "[dry-run] Would run: node $REPO_ROOT/dist/index.js auth"
      SIDECAR_REAUTH_PENDING=false
    else
      # Auth flow overwrites the credentials file on success, so we don't
      # delete the existing token first — keeps it as a fallback if the
      # browser flow is cancelled or fails.
      if node "$REPO_ROOT/dist/index.js" auth; then
        success "Re-authentication complete!"
        SIDECAR_REAUTH_PENDING=false
      else
        error "Re-authentication failed. Your existing credentials are unchanged."
        SIDECAR_REAUTH_PENDING=true
        info "You can retry later with: ./scripts/upgrade.sh"
      fi
    fi
  else
    SIDECAR_REAUTH_PENDING=true
    info "Skipped. To re-authenticate later, run: ./scripts/upgrade.sh"
  fi
}

main
