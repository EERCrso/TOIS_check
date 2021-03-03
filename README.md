# About
This is a work in progress, aimed for more efficient reporting of internal issues.

## Currently implemented inputs
- Atlassian JIRA .csv complete export
- HP Service Manager .xlsx export
- previous month report in .xlsx format

## Currently implemented outputs
- Long and short .xlsx reports
- JIRA check of created incidents in last month

## Usage
- needs to run in Python (developing on Python 3.8) terminal or IDE
- GUI interface will guide you through inputs and outputs

## Files
- main_merge.py - main script file
- aux_functions.py - auxiliary functions for the main file, must be in the same directory
- tois_check_main.py - discontinued branch (loading data from .xml files)
