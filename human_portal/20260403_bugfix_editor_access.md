# Bug Fix Record: Dashboard Editor Access Denied
Date: 2026-04-03
Time: 14:45
Issue: When clicking "View/Edit (New Window)" in the dashboard, an "Access Denied" alert was shown despite being logged in.
Root Cause: `editor.html` was checking for `localStorage.getItem('userGrade')`, but the login logic in `script.js` only set `localStorage.setItem('userSession', ...)`.
Resolution: Updated `editor.html` to check for `userSession` in `localStorage` instead of `userGrade`.
Files Modified: 
- editor.html
