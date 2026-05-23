#!/bin/bash
# Run this once from inside the email-guard folder on your machine

git init
git add .
git commit -m "Initial commit — Email Guard add-in"
git branch -M main
git remote add origin https://github.com/jameswong95/email-guard.git
git push -u origin main
