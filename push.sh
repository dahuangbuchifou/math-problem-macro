#!/bin/bash
cd /home/admin/.openclaw/workspace/math-problem-macro
LOG="/home/admin/.openclaw/workspace/math-problem-macro/push.log"
echo "$(date '+%Y-%m-%d %H:%M:%S') - Attempting push..." >> "$LOG"
timeout 60 git push origin main >> "$LOG" 2>&1
if [ $? -eq 0 ]; then
    echo "$(date '+%Y-%m-%d %H:%M:%S') - Push SUCCESS!" >> "$LOG"
    # Stop cron after success
    crontab -l | grep -v "push.sh" | crontab -
    echo "Push successful! Cron stopped."
else
    echo "$(date '+%Y-%m-%d %H:%M:%S') - Push failed (exit code $?). Will retry." >> "$LOG"
fi
