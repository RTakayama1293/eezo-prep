#!/bin/bash
# リモート環境（Claude Code）でのみ実行
if [ "$CLAUDE_CODE_REMOTE" != "true" ]; then
  echo "Skipping setup (not in Claude Code environment)"
  exit 0
fi
echo "Setting up EEZO-prep environment..."
pip install -r requirements.txt
echo "Setup complete!"
exit 0
