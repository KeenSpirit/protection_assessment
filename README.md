ProtectionAssessment
===============

Code under development on Gitea at the following location:
http://smartweb.myergon.local:3000/dp072/protection-assessment

Creates the following files under the active study case for a given selection of relays:
- Protection section time overcurrent plots
- Protection section single line diagram*
- Protection section audit

Map the start.py file to PowerFactory script as the executable.

* A single line diagram is only created if a template file is stored at the following location within the PowerFactory Data Manager:
/user/templates
* A copy of this template file is stored in the following directory of the codebase:
* /pf SLD template file