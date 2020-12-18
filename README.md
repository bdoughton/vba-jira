VBA-Jira
=======

VBA-Jira is designed to make it easy for Jira users on Windows and Mac to track progress of projects and teams via MS Excel. It includes support REST API calls using <a href="https://github.com/VBA-tools/VBA-Web">VBA-Web</a> by Tim Hall et al.

Getting started
---------------

- Download the [latest release (v1)](https://github.com/bdoughton/vba-jira/releases)
- The project is composed of an MS Excel Add-in: `vba-jira.xlam` with a CustomUI Ribbon and associated excel templates:
    - Metrics: `ScrumTeamStats.xlsx`

Notes
---

Modules are structred as follows:
- `Jira.bas`: Common Jira functions 
- `JiraAgile.bas`: Functions for exporting/updating Jira Agile boards to the active worksheet
- `JiraScrumTeamStats.bas`: Functions for compiling metrics (see associated excel template)
- `JiraAffinityEstimation.bas`: Functions for exporting/udating Story Porint estimates for Jira issues

Before commiting changes to Git the source code should be exported from the Visual Basic editor to the src directory. The following <a href="https://www.xltrail.com/blog/auto-export-vba-commit-hook">post</a> from Björn Stiel explains some useful options for automating this.

### Release Notes

View the [changelog](https://github.com/bdoughton/vba-jira/CHANGELOG.md) for release notes

### About

- Author: Ben Doughton
- License: GNU General Public License v3.0
