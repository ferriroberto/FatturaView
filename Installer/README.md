FatturaView Installer helper

This folder contains artifacts to build a Windows installer (Inno Setup) and helper script.

Steps to create a release installer on your machine:

1. Build Release binaries
   - Open "Developer PowerShell for Visual Studio" and run:
     .\build_release.ps1
   - This will build the solution (Release x64) and copy executables and resources into `installer\dist`.

2. Create installer
   - Install Inno Setup (https://jrsoftware.org/isinfo.php)
   - From Inno Setup, open `installer\setup.iss` and compile, or run:
       ISCC.exe installer\setup.iss

Notes:
- The installer shows the LICENSE file during installation (license acceptance).
- The installer includes everything inside `installer\dist` (binaries and LICENSE). Ensure any required DLLs are copied there by the build script.
