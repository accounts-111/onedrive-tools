#!/bin/bash
# Double-click this file to launch OneDrive Tools

cd "$(dirname "$0")"
echo "Starting OneDrive Tools..."
echo ""
python3 onedrive_tools.py
echo ""
echo "Press any key to close..."
read -n 1
