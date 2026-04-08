@echo off
cd /d C:\GITHUB\anketa_web

echo ADD
git add .

echo COMMIT
git commit -m "auto update"

echo PUSH
git push origin main

exit