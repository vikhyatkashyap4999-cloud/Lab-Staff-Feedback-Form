# Start here (on the VM)

1. **Unzip** this file into `E:\`. You should get `E:\labfeedback\` with `app.py` and `web.config` inside.
   *(If you use another folder, change the 2 paths marked "CHANGE IF NEEDED" in `web.config`.)*
2. **Keep real Excel files outside this folder**, for example in `E:\LabData\`. Upload them later from the admin page.
3. **Check the server**: open PowerShell as Administrator in this folder and run:
   ```
   powershell -ExecutionPolicy Bypass -File .\check_server.ps1
   ```
4. **Follow `HOSTING_GUIDE_IIS.md`**, step by step, until the checker shows everything OK.
5. **Open this folder in Antigravity** and make your first message:
   > Read AGENTS.md and the project files, then tell me what this project is and what is pending.
