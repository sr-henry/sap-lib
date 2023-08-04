# SAP Automation Modules

SAP Automation Modules is a collection of simple modules designed to make it easy to interact with SAP GUI Scripting API and SAP Analysis for Office. These modules provide a straightforward and user-friendly way to automate various tasks in SAP, improving efficiency and reducing manual efforts.

## Table of Contents
- [Introduction](#introduction)
- [Documentation](#documentation)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Examples](#examples)

## Introduction
SAP GUI Scripting API and SAP Analysis for Office are powerful tools for automating SAP processes and performing data analysis. However, interacting with these APIs can be complex and time-consuming, especially for users without extensive programming experience. This repo aims to bridge this gap by providing easy-to-use modules that abstract away the complexities of the underlying APIs, making automation tasks more accessible to everyone.

## Documentation
- [SAP GUI Scripting API](https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.00/en-US)
- [SAP Analysis for Microsoft Office](https://help.sap.com/docs/SAP_BUSINESSOBJECTS_ANALYSIS_OFFICE/ca9c58444d64420d99d6c136a3207632/ebf198667aa54740b9049d9da804a901.html?version=2.8.8.0)

## Features
- **User-friendly Interface:** The modules offer a simplified and intuitive interface for interacting with SAP GUI and SAP Analysis for Office, reducing the learning curve.
- **Easy Installation:** Installation is straightforward and can be done with just a few simple steps.
- **Customization:** The modules can be easily customized to fit specific business requirements and workflows.

## Requirements
- SAP GUI Scripting API enabled on your SAP system.
- SAP Analysis for Office installed and configured.

## Installation
To install SAP Automation Modules, follow these steps:

1. Clone this repository to your local machine or download the latest release.
2. Ensure you have the required dependencies installed (SAP GUI Scripting API and SAP Analysis for Office).
3. Copy the module files into your project directory or include them in your automation workflow.

## Usage
To use SAP Automation Modules in your project, follow the documentation and examples provided. Each module has its own set of functions and methods to interact with SAP GUI and SAP Analysis for Office.

### SAP GUI Scripting

For example, to automate a SAP transaction using the `py` module:

```python
import sap

sap_connection_data = sap.attach("system name")

if sap_connection_data:

    application, connection, session = sap_connection_data

    # script here
    session.StartTransaction("tcode")
 
    sap.close(sap_connection_data)
```
And using the `vba` class module:
```VBA
Dim sap As New SapGuiScripting
If sap.Attach("system name") Then

    ' script here
    session.StartTransaction "tcode"
 
End If
```
### SAP Analysis for Office

To automate a refresh of SAP Analysis report using the `py` module:
```python
from ao import SapAnalysisOffice

with SapAnalysisOffice('Ao_Workbook_File_Path') as ao:

   if not 'DS_NAME' in ao.datasources:
       print(f'Data Source "DS_NAME" does not exist in Workbook "{ao.wb.Name}"')

   if not ao.is_connected(data_src):
       ao.refresh(data_src)
```
Using `vba` class module:
```VBA
Dim prompts As New Dictionary
With prompts
    .Add "<variable technical name 0>", "<variable value 0>"
    .Add "<variable technical name 1>", "<variable value 1>"
    .Add "<variable technical name 2>", "<variable value 2>"
End With

Dim ao As New SapAnalysisOffice
If Not ao.Refresh("DS_1", prompts) Then
    Debug.Print "Refresh AO fail!"
End If
```

## Examples
Check out the [Examples](examples/) directory for more detailed use cases and practical examples of how to leverage SAP Automation Modules for various automation tasks.
