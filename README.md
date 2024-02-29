# ffpreader

## Introduction

Ampac Configuration Manager PLUS Is a software tool designed and written by AMPAC Technologies Pty. Ltd to enable users to create configuration files for transfer to and from the FireFinder PLUS Fire Alarm Control Panel.

As it stands, Configuration Manager PLUS is not equipped with a feature to export data in a human friendly format, which makes 3rd party integration and change management very difficult.

The purpose of this tool is to read Ampac FireFinder PLUS *.ffp Configuration Files, translate into a human-readable format and write to csv and Excel workbooks.

An additional function of the tool is to track changes (work in progress), so that multiple revisions of a program can be compared and changes summarised, eg a new detector is added, a device is moved to a new zone.

## Usage

Requires python v3 to run.

The functions and working code is currently contained in `ffpreader.py`. There is an example input .ffp file and example output csv files and Excel workbook.

The following types of equipment are tabulated in output files:

- devices
- loops
- nodes 
- zones

For Modbus integration, these lists are split out into separte Excel sheets for the Modbus register mapping, eg loops 1-90, 91-180, 181-250.
