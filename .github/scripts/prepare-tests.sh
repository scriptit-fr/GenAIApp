#!/usr/bin/env bash
set -euo pipefail

TEST_FILE="src/testFunctions.gs"
ENABLE_API_KEY_TESTS="${ENABLE_API_KEY_TESTS:-false}"
ENABLE_VERTEX_AI_TESTS="${ENABLE_VERTEX_AI_TESTS:-false}"
OPEN_AI_API_KEY="${OPEN_AI_API_KEY:-}"
GEMINI_API_KEY="${GEMINI_API_KEY:-}"
VERTEX_AI_PROJECT_ID="${VERTEX_AI_PROJECT_ID:-}"
VERTEX_AI_LOCATION="${VERTEX_AI_LOCATION:-global}"

if [ ! -f "$TEST_FILE" ]; then
  echo "::error::${TEST_FILE} not found; cannot prepare remote Apps Script tests."
  exit 1
fi

if [ "$ENABLE_API_KEY_TESTS" = "true" ]; then
  missing=()
  if [ -z "$OPEN_AI_API_KEY" ]; then
    missing+=("OPEN_AI_API_KEY")
  fi
  if [ -z "$GEMINI_API_KEY" ]; then
    missing+=("GEMINI_API_KEY")
  fi

  if [ "${#missing[@]}" -gt 0 ]; then
    printf '::error::Cannot prepare API key tests because required environment variable(s) are missing: %s\n' "${missing[*]}"
    exit 1
  fi
else
  echo "Skipping API key validation: API key tests are disabled."
fi

if [ "$ENABLE_VERTEX_AI_TESTS" = "true" ]; then
  if [ -z "$VERTEX_AI_PROJECT_ID" ]; then
    echo "::error::Cannot prepare Vertex AI tests because VERTEX_AI_PROJECT_ID is missing."
    exit 1
  fi
else
  echo "Skipping Vertex AI configuration validation: Vertex AI tests are disabled."
fi

echo "Validated remote Apps Script test configuration without modifying ${TEST_FILE}."
