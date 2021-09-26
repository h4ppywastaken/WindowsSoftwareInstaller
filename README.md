# WindowsSoftwareInstaller
Powershell script to install a list of software on the local computer automatically

## Software Installer
Powershell Script to install a list of software on the local computer configurable through a configuration file.
This script is supposed to help with tasks like reinstalling Windows.

Supported Methods:
   - Direct Download Url
   - Sourceforge ("latest" link)
   - Github
   - Chocolatey
   - Manual User Download (just opens the link from the config in firefox browser)

See WSISW.config for examples.

This script will automatically install chocolatey.

It can also create a Windows Scheduled Task to automatically check for chocolatey package updates on each system startup. You can set exclusions to this automatic upgrade or skip creating this task.

If this script is not run as admin it will automatically elevate privileges.

# License
This project is licensed under the MIT License. See LICENSE file for details.
