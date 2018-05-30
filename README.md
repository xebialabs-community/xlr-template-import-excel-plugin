# XL Release Template Import Excel plugin

[![Build Status][xlr-template-import-excel-plugin-travis-image]][xlr-template-import-excel-plugin-travis-url]
[![License: MIT][xlr-template-import-excel-plugin-license-image]][xlr-template-import-excel-plugin-license-url]
![Github All Releases][xlr-template-import-excel-plugin-downloads-image]

[xlr-template-import-excel-plugin-travis-image]: https://travis-ci.org/xebialabs-community/xlr-template-import-excel-plugin.svg?branch=master
[xlr-template-import-excel-plugin-travis-url]: https://travis-ci.org/xebialabs-community/xlr-template-import-excel-plugin
[xlr-template-import-excel-plugin-license-image]: https://img.shields.io/badge/License-MIT-yellow.svg
[xlr-template-import-excel-plugin-license-url]: https://opensource.org/licenses/MIT
[xlr-template-import-excel-plugin-downloads-image]: https://img.shields.io/github/downloads/xebialabs-community/xlr-template-import-excel-plugin/total.svg

## Preface

This document describes the functionality provided by the XL Release Template Import Excel plugin.

See the [XL Release reference manual](https://docs.xebialabs.com/xl-release) for background information on XL Release and release automation concepts.  

## Overview

This plugin will convert an Excel spreadsheet into an XL Release template.  The Import-from-Attachments task will import any attached spreadsheets.  Or, a spreadsheet file can be passed via the XL Release REST API.

See the samples directory for an example of the formatting.  Type, User, Team, Duration, and Description columns are optional.

Coding for Team and Duration is not complete, so those columns will be ignored pending additional work.

## Requirements

* XL Release 7+

## Installation

* Copy the latest JAR file from the [releases page](https://github.com/xebialabs-community/xlr-template-import-excel-plugin/releases) into the `XL_RELEASE_SERVER/plugins` directory.
* Restart the XL Release server.

## Features/Usage/Types/Tasks

* The Import-from-Attachments task imports multiple Excel spreadsheet attachments on a task and converts then into templates base on the file name.
* The plugin can also accept an Excel spreadsheet file via the REST API.

![[import-from-attachments-task]](images/import-from-attachments-task.png)

## References

[Apache POI-HSSF/POI-XSSF](https://poi.apache.org/spreadsheet/)

