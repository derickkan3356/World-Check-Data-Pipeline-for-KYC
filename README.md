# World Check Data Pipeline for KYC
This repository showcases the development of a data pipeline and storage solution for compliance team, to handle World Check KYC screening data through Excel VBA and SQL Server integration. It's designed to streamline the process of uploading investor information for screening and storing results efficiently.

## Project Overview
The project consists of a VBA module, SQL queries, a custom Excel UI, and several Excel templates, aimed at automating the import and export of data between Excel and a SQL Server database. This automation supports the KYC compliance workflow, including data submission to third-party screening platforms and the internal consolidation of screening results.

### Components
**World check data pipeline.bas**: A VBA module containing key functions to interact with the SQL Server database, including data import/export and template generation.

**SQL query.sql**: SQL scripts for creating necessary tables and procedures in the SQL Server database.

**customUI.xml**: XML definition for a custom Excel ribbon UI, facilitating easy access to the VBA functions.

**Button callback.bas**: VBA code connecting the custom UI buttons with their respective VBA functions.
![Ribbon button](https://github.com/derickkan3356/World-Check-Data-Pipeline-for-KYC/blob/main/ribbon%20button.png)

**Excel Templates**: Located in the Template folder, these Excel files serve as templates for uploading data to SQL Server, extracting data for third-party screening, and formatting screening results for re-import.

### Key Functions
**ImportWCdata()**: Uploads data from Excel to the SQL Server database.

**ExportWCdata()**: Prepares investor data in a format ready for uploading to a third-party name screening platform.

**ExportWCtemplate()**: Generates an Excel template populated with data from SQL Server for screening result entry.

## Usage
This repository is intended as a demonstration of my contributions to the project, focusing on VBA-SQL Server integration and automation within Excel. It does not include the complete application or external dependencies, such as user-defined functions (UDFs) and SQL server tables developed by other team members.

## Disclaimer
The code provided in this repository is for demonstration purposes only and is part of a larger compliance project. It may require modifications to work in your environment or to connect with other components not included here.