ProtectionAssessment
===============

Code under development on Gitea at the following location:
http://smartweb.myergon.local:3000/dp072/protection-assessment

Performs a feeder fault study and conductor damage assesment and saves the results to the user's home path.

Creates the following files under the active study case for a given selection of relays:
- Protection section time overcurrent plots
- Protection section single line diagram*
- Protection section audit

A conductor damange assessment is visible in the graphical interface by setting feeder colour scheme to User-Defined/{fault type}.

All protection devices (elmrelay and elmfuse) must be correctly connected and configured for the script to work correctly.

Map the start.py file to PowerFactory script as the executable.

* This feature is currently disabled