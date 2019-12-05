<p align="center">
  <img src="https://github.com/abasile-ml/testeo-git/blob/master/img/logo.png" alt="MELI" width="500">
</p>
<p align="center">An API to export S3 data to google sheet.</p>
<hr>
<p align="center">
 <img src="https://img.shields.io/badge/team%20-%20shipping--metrics-green">
<a href="https://app.intercom.io/a/apps/avw9yqcm/home"><img src="https://img.shields.io/badge/support%20-%20rmansilla-blue"></a>
<img src="https://img.shields.io/badge/java%20-%20v1.8-orange">
</p>

**Gaia Gsheet export API** aims to help users quickly export data from S3 to a google sheet. To make easy to the user the data reading, an style is applyed to the sheet and a pivot table is created with the same purpose.

<p align="center">
<img src="https://github.com/abasile-ml/testeo-git/blob/master/img/data.png"
width="350">
<img src="https://github.com/abasile-ml/testeo-git/blob/master/img/pivot-table.png"
width="350">
</p>

## Table of Contents

* [Getting started](#getting-started)
  * [Dependencies](#dependencies)
  * [Configuration](#configuration)
* [Uses](#uses)
  * [Limitations](#limitations)
  
## Getting started
```shell
$ fury get gaia-gsheet-export
```
  
## Uses
-Export information from S3 to Google sheet.
-Google sheet apply format keeping the data type.

## Limitations
5000000 rows allow(?)
800000 values.
