# SoftwareMatrix
A series of scripts to document and report on software changes

Requires MySQL 8.0
Blank database is included. Plan is to create a script to create database and prompt for needed info.

Scripts and their usage:

IngestCSV.vbs:
This is the main script. Frequency should be daily.
Pulls in a CSV file in the format (Workstation, Application, Publisher, Version) from any software collection source (we use PDQ - www.pdq.com.)
Checks any new software to see if it is FOSS using two websites (www.fosshub.com and www.chocolatey.org)
Imports collected data into database and then reports on changes.
Changes are reported via two emails - the Security Report and the Change Report
- Security Report only reports on changes that impact the organization as a whole (e.g. new software added or deleted that was never seen before.) Only gets generated when changes dictate.
- Change Report includes all changes since the last import for all PCs.

CheckSoftwareUpdates.vbs:
Checks each installed application to see if version matches the version on the provided website when it was categorized. Frequency should be weekly.
This unique way of checking for updates is not fool proof. There are a lot of false positives but these can be minimized by the use of the variance field.
An email report is generated.

CheckVulnerabilities.vbs:
Checks nvd.nist.gov for vulnerabilities in installed applications. Frequency should be weekly (important because feed only keeps about a week of data.)
Uses RSS feed on NIST website to generate an email report on any applications thats name matches the feed. Not 100% accurate. Name must be exactly the same to match.
Will report on a version match if there is one.

CountApps-AssociatedReports.vbs
Counts total installed instances per application. Frequency - monthly.
Compares count to licenses table. Right now data needs to be entered to that table manually.
Creates email report if any applications are over subscribed.
Creates email report on top 10 highest risk software based on vulnerabilities within last year.

EnterAppDetails.vbs:
Run as needed.
Allows the admin user to quickly answer questions about newly found applications (from the Security Report) to get all of the columns filled in for the application.
The hope is to replace this with a web GUI.


To Do:
- Create web GUI - Project started - https://github.com/compuvin/SoftwareMatrix-GUI
- Create and Integrate WhatTheFOSS list
- Better code documentation
