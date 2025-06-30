#!/usr/bin/env bash

# Install Chromium needed by Puppeteer
apt-get update && apt-get install -y chromium

# Optional: Set Puppeteer to use this Chromium
export PUPPETEER_EXECUTABLE_PATH=$(which chromium)
