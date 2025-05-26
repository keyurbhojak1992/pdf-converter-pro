#!/usr/bin/env bash
echo "-----> Checking for LibreOffice"
if ! command -v libreoffice &> /dev/null; then
    echo "-----> Installing LibreOffice via apt"
    apt-get update -y
    apt-get install -y libreoffice
else
    echo "-----> LibreOffice already installed"
fi
