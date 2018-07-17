---
title: Application.FileSaveAs Method (Project)
keywords: vbapj.chm107
f1_keywords:
- vbapj.chm107
ms.prod: project-server
api_name:
- Project.Application.FileSaveAs
ms.assetid: 0b5fe86c-28ea-5a9e-53df-5a83030c0d20
ms.date: 06/08/2017
---


# Application.FileSaveAs Method (Project)

Saves the active project to a new file name or exports data to a file.


## Syntax

_expression_. **FileSaveAs** (**_Name_**, **_Format_**, **_Backup_**, **_ReadOnly_**, **_TaskInformation_**, **_Filtered_**, **_Table_**, **_UserID_**, **_DatabasePassWord_**, **_FormatID_**, **_Map_**, **_Password_**, **_WriteResPassword_**, **_ClearBaseline_**, **_ClearActuals_**, **_ClearResourceRates_**, **_ClearFixedCosts_**, **_XMLName_**, **_ClearConfirmed_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of a project file.|
| _Format_|Optional|**PjFileFormat**|Specifies the file format. The format of the file. The _FormatID_ argument should be used in place of _Format_, which is included primarily for backwards compatibility. If _FormatID_ is specified, _Format_ is ignored. Can be one of the **[PjFileFormat](pjfileformat-enumeration-project.md)** constants. The default is **pjMPP**.|
| _Backup_|Optional|**Boolean**|**True** if Project creates a backup copy of the file. The default is **False**.|
| _ReadOnly_|Optional|**Boolean**|**True** if Project should display an alert recommending that the file be opened read-only. If selectively exporting data instead of saving a complete project, _ReadOnly_ is ignored. The default value is **False**.|
| _TaskInformation_|Optional|**Boolean**|**True** if task information is saved in a non-project file format. The _Map_ argument should be used in place of _TaskInformation_, which is included primarily for backward compatibility. If _Map_ is specified, _TaskInformation_ is ignored. The default value is **True** if the active view is a task view; otherwise it is **False**.|
| _Filtered_|Optional|**Boolean**|**True** if filtered tasks or resources are saved, for a project saved in a non-Project file format. **False** if all the tasks or resources are saved. If _Map_ is specified, _Filtered_ is ignored. The default value is **False**.|
| _Table_|Optional|**Variant**|The name of the table containing resource or task information for a project saved in a non-Project file format. The _Map_ argument should be used in place of _Table_, which is included for backward compatibility. If _Map_ is specified, or _Name_ specifies a project file format, _Table_ is ignored. The default value is the name of the active table.|
| _UserID_|Optional|**String**|Not used. Project can open a project file that an earlier version of Project saved to an ODBC database, but cannot save to a database.|
| _DatabasePassWord_|Optional|**String**|Not used. Project cannot save to an ODBC database.|
| _FormatID_|Optional|**String**|Specifies the file format to use. If Project recognizes the format of the file specified by _Name_,  _FormatID_ is ignored. _FormatID_ can be one of the [following format string values](#format-strings) for saving files.|
| _Map_|Optional|**String**|The name of the import/export map to use when exporting data.|
| _Password_|Optional|**String**|A password to use when opening password-protected project files. If _Password_ is incorrect or omitted and a file requires a password, the user is prompted for the password.|
| _WriteResPassword_|Optional|**String**|A password to use when writing to a write-reserved project file. If _WriteResPassword_ is omitted and the file requires a password, the user is prompted for the password.|
| _ClearBaseline_|Optional|**Boolean**|**True** if baseline values (the Baseline Cost, Baseline Work, Baseline Start, Baseline Finish, Baseline Duration, Timephased Baseline Work, and Timephased Baseline Cost fields) are cleared when saving as a template. The default value is **False**.|
| _ClearActuals_|Optional|**Boolean**|**True** if actual values (the % Complete field and, if actual costs are not calculated by Project, the Actual Cost field) are cleared when saving as a template. The default value is **False**.|
| _ClearResourceRates_|Optional|**Boolean**|**True** if resource rate tables are cleared when saving as a template. The default value is **False**.|
| _ClearFixedCosts_|Optional|**Boolean**|**True** if the Fixed Costs field is cleared for all tasks when saving as a template. The default value is **False**.|
| _XMLName_|Optional|**Variant**|This is the XML DOM object that is passed to the function when _FormatID_ is "MSProject.XML". The **FileSaveAs** method fails if the XML format is specified and _XMLName_ is not a valid XML DOM object. If _FormatID_ is anything other than "MSProject.XML", _XMLName_ should be **NULL** and the method should fail. Only one of _XMLName_ or _Name_ can be specified.|
| _ClearConfirmed_|Optional|**Boolean**|**True** if the information is cleared about whether tasks have been confirmed as published to Project Server. The default value is **False**.|

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


### Return value

 **Boolean**

## Remarks

Using the value "MSProject.mpp.9" for the _FormatID_ parameter causes Project to show the **Saving to Previous Version - Compatibility Checker** dialog box. For example, manually scheduled tasks will be converted to automatically scheduled tasks in previous Project versions. You can choose to keep the format or cancel the save operation. You can also select **Don't tell me about this again**.

> [!NOTE]
> Several _FormatID_ strings are obsolete; if you try to use them, they result in run-time error 1004. _FormatID_ values such as "MSProject.odbc" can be used in Project 2003 and earlier versions but are removed in Project 2007 and later versions.


