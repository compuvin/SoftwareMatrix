# SoftwareMatrix
A series of scripts to document and report on software changes

<!-- GETTING STARTED -->
## Getting Started

The goal is to make the process of getting up and running quick and easy. Follow the instructions below and you'll be tracking your installed software in no time!

There is also a web based GUI that is available - https://github.com/compuvin/SoftwareMatrix-GUI

### Prerequisites

Requires MySQL 8.0
Script will prompt for server information and create the necessary tables.

### Installation

To install, simply clone the repo and schedule tasks based on the information in the <a href="#usage">usage</a> section.

1. Clone the repo
   ```sh
   git clone https://github.com/compuvin/SoftwareMatrix.git
   ```
2. Setup Scheduled Tasks (see <a href="#usage">usage</a>)

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- USAGE EXAMPLES -->
## Usage

Scripts and their usage:

IngestCSV.vbs:
This is the main script. Frequency should be daily.
Pulls in a CSV file in the format (Workstation, Application, Publisher, Version) from any software collection source (we use PDQ - www.pdq.com.)
Checks any new software to see if it is FOSS using two websites (www.openhub.net and www.chocolatey.org)
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
This has been replaced with the web GUI but can still be used.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ROADMAP -->
## Roadmap

- [ ] Create a software crawler or agent for small deployments
- [ ] Better code documentation
- [ ] Create and Integrate WhatTheFOSS list

See the [open issues](https://github.com/compuvin/SoftwareMatrix/issues) for a full list of proposed features (and known issues).

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- LICENSE -->
## License

Distributed under the GPL-3.0 license. See `LICENSE` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>
