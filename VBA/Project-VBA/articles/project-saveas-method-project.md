---
title: Project.SaveAs Method (Project)
keywords: vbapj.chm132597
f1_keywords:
- vbapj.chm132597
ms.prod: project-server
api_name:
- Project.Project.SaveAs
ms.assetid: 947fb1f9-0abd-7423-2c22-96bb91f2dc6e
ms.date: 06/08/2017
---


# Project.SaveAs Method (Project)

Saves a file that is not the active project under a new file name.

## Syntax

_expression_. **SaveAs** (**_Name_**, **_Format_**, **_Backup_**, **_ReadOnly_**, **_TaskInformation_**, **_Filtered_**, **_Table_**, **_UserID_**, **_DatabasePassWord_**, **_FormatID_**, **_Map_**, **_ClearBaseline_**, **_ClearActuals_**, **_ClearResourceRates_**, **_ClearFixedCosts_**)

_expression_ A variable that represents a **Project** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the file to save as. If _Name_ is omitted, Project prompts for the file name.|
| _Format_|Optional|**Long**|The format of the file. The _FormatID_ argument should be used in place of _Format_, which is included primarily for backwards compatibility. If _FormatID_ is specified, _Format_ is ignored. The default value is **pjMPP**. Can be one of the **[PjFileFormat](pjfileformat-enumeration-project.md)** constants.|
| _Backup_|Optional|**Boolean**|**True** if Project makes a backup copy of the file.|
| _ReadOnly_|Optional|**Boolean**|**True** if Project should display an alert recommending that the project be opened read-only. The default value is **False**.|
| _TaskInformation_|Optional|**Boolean**|**True** if task information is saved, for a project saved under a non-Project file format. **False** if resource information is saved. If _Map_ is specified, _TaskInformation_ is ignored. The default value is **True** if the active view is a task view, and **False** otherwise.|
| _Filtered_|Optional|**Boolean**|**True** if filtered tasks or resources are saved, for a project saved under a non-Project file format. **False** if all the tasks or resources are saved. If _Map_ is specified, _Filtered_ is ignored. The default value is **False**.|
| _Table_|Optional|**String**|The name of the table containing the task or resource information, for a project saved under a non-Project format. If _Map_ is specified, or _Name_ specifies a database file or format, _Table_ is ignored. The default value is the name of the active table.|
| _UserID_|Optional|**String**|Not used. Project can open a project file that an earlier version of Project saved to an ODBC database, but cannot save to an ODBC database.|
| _DatabasePassWord_|Optional|**String**|Not used. Project cannot save to an ODBC database.|
| _FormatID_|Optional|**String**|Specifies the file or database format to use. If Project recognizes the format of the file specified with _Name_, _FormatID_ is ignored. _FormatID_ can be one of the values in the [Format strings](#format-strings) table.|
| _Map_|Optional|**String**|The name of the import/export map to use when exporting data.|
| _ClearBaseline_|Optional|**Boolean**|**True** if baseline values (the Baseline Cost, Baseline Work, Baseline Start, Baseline Finish, Baseline Duration, Timephased Baseline Work, and Timephased Baseline Cost fields) are cleared when saving as a template. The default value is **False**.|
| _ClearActuals_|Optional|**Boolean**|**True** if actual values (the % Complete field and, if actual costs are not calculated by Project, the Actual Cost field) are cleared when saving as a template. The default value is **False**.|
| _ClearResourceRates_|Optional|**Boolean**|**True** if resource rate tables are cleared when saving as a template. The default value is **False**.|
| _ClearFixedCosts_|Optional|**Boolean**|**True** if the Fixed Costs field is cleared for all tasks when saving as a template. The default value is **False**.|

<br/>

#### Format strings

|**Format string**|**Description**|
|:-----|:-----|
|"MSProject.mpp"|Project file|
|"MSProject.mpt"|Project template|
|"MSProject.mpp.8"|Project 98 file|
|"MSProject.mpp.9"|Project 2000–Project 2003 file|
|"MSProject.mpp.12"|Project 2007 file|
|"MSProject.xls"|Excel workbook|
|"MSProject.xls5"|Excel 97–Excel 2003 workbook|
|"MSProject.pdf"|PDF file|
|"MSProject.xpf"|XPF file|
|"MSProject.csv"|CSV (comma delimited) file|
|"MSProject.txt"|TXT (tab delimited) file|
|"MSProject.xml"|Project XML file|

<br/>

## Remarks

Using the value "MSProject.mpp.9" for the _FormatID_ parameter causes Project to show the **Saving to Previous Version - Compatibility Checker** dialog box. For example, manually scheduled tasks will be converted to automatically scheduled tasks in previous Project versions. You can choose to keep the format or cancel the save operation. You can also select **Don't tell me about this again**.

> [!NOTE]
> Several _FormatID_ strings are obsolete; if you try to use them, they result in run-time error 1004. _FormatID_ values such as "MSProject.odbc" can be used in Project 2003 and earlier versions but are removed in Project 2007 and later versions.

