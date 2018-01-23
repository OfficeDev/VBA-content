---
title: Application.FileOpenEx Method (Project)
keywords: vbapj.chm102
f1_keywords:
- vbapj.chm102
ms.prod: project-server
api_name:
- Project.Application.FileOpenEx
ms.assetid: d03c13b0-c12f-1d45-bb80-26711d69a378
ms.date: 06/08/2017
---


# Application.FileOpenEx Method (Project)

Opens a project or imports data.


## Syntax

_expression_. **FileOpenEx** (**_Name_**, **_ReadOnly_**, **_Merge_**, **_TaskInformation_**, **_Table_**, **_Sheet_**, **_NoAuto_**, **_UserID_**, **_DatabasePassWord_**, **_FormatID_**, **_Map_**, **_openPool_**, **_Password_**, **_WriteResPassword_**, **_IgnoreReadOnlyRecommended_**, **_XMLName_**, **_DoNotLoadFromEnterprise_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the project file, source file, or data source to open. If _Name_ is not specified, Project displays the **Open** dialog box.|
| _ReadOnly_|Optional|**Boolean**|**True** if the file is opened read-only. If selectively importing data instead of loading a complete project, _ReadOnly_ is ignored.|
| _Merge_|Optional|**Long**|Specifies whether to automatically merge the file (MPX and XMLDOM formats only) with the active project. To automatically merge XLS, CSV, or TXT file formats, you can set the merge key in the import map that you are using. The _Map_ argument should be used in place of _Merge_, which is included for backward compatibility. If _Map_ is specified, _Merge_ is ignored. Can be one of the [PjMergeType](pjmergetype-enumeration-project.md) constants. The default value is **pjDoNotMerge**.|
| _TaskInformation_|Optional|**Boolean**|**True** if the file contains information about tasks for a project saved under a non-Project file format. **False** if the file contains information about resources. The _Map_ argument should be used in place of _TaskInformation_, which is included for backward compatibility. If _Map_ is specified, _TaskInformation_ is ignored. The default value is **True** if the active view is a task view; otherwise it is **False**.|
| _Table_|Optional|**String**|The name of a table in which to place the resource or task information for a project saved under a non-Project file format.  _Table_ is required if the value of the _Merge_ argument is **pjMerge**. The _Map_ argument should be used in place of _Table_, which is included for backward compatibility. If _Map_ is specified, or _Name_ specifies a database file or format, _Table_ is ignored. The default value for _Table_ is the name of the active table.|
| _Sheet_|Optional|**String**|The sheet to read when opening a workbook created in Excel version 5.0 or later. The _Map_ argument should be used in place of _Sheet_, which is included for backward compatibility. If _Map_ is specified, or if the file specified by _Name_ is not an Excel file, _Sheet_ is ignored.|
| _NoAuto_|Optional|**Boolean**|**True** if any **Auto_Open** macro is prevented from running. The default value is **False**.|
| _UserID_|Optional|**String**|A user ID to use when accessing a database. If _Name_ or _FormatID_ is not a database, _UserID_ is ignored.|
| _DatabasePassWord_|Optional|**String**|A password to use when accessing a database. If _Name_ or _FormatID_ is not a database, _DatabasePassWord_ is ignored.|
| _FormatID_|Optional|**String**|Specifies the file or database format to use. If Project recognizes the format of the file specified with _Name_, _FormatID_ is ignored. _FormatID_ can be one of the values in the [Format strings](#format-strings) table.|
| _Map_|Optional|**String**|The name of the import/export map to use when importing data.|
| _openPool_|Optional|**Long**|The action to take when opening a resource pool or sharer file. When opening a master project, the value for _openPool_ is also applied to the subprojects. Can be one of the [PjPoolOpen](pjpoolopen-enumeration-project.md) constants. The default value is **pjPromptPool**.|
| _Password_|Optional|**String**|A password to use when opening password-protected project files. If _Password_ is incorrect or omitted and a file requires a password, the user is prompted for the password.|
| _WriteResPassword_|Optional|**String**|A password to use when writing to a write-reserved project file. If _WriteResPassword_ is omitted and the file requires a password, the user is prompted for the password.|
| _IgnoreReadOnlyRecommended_|Optional|**BooleanVariant**|**True** to prevent Project from displaying an alert that the project should be opened read-only. If the project was not saved with a read-only recommendation, _IgnoreReadOnlyRecommended_ is ignored.|
| _XMLName_|Optional|**Variant**|This is the XML DOM object that is passed to the function when _FormatID_ is MSProject.XML. The **FileSaveAs** method fails if the XML format is specified and _XMLName_ is not a valid XML DOM object. If _FormatID_ is anything other than MSProject.XML, _XMLName_ should be **NULL** and the method should fail otherwise. Only one of _XMLName_ or _Name_ can be specified.|
| _DoNotLoadFromEnterprise_|Optional|**Boolean**|**True**, if the project is not to be opened from Project Server. The default is **False**, where Project Professional opens the file from Project Server, or from the local computer if Project Professional is not logged on Project Server.|

<br/>

#### Format strings

|**Format string**|**Description**|
|:-----|:-----|
|"MSProject.mpp"|Project file|
|"MSProject.mpt"|Project template|
|"MSProject.mpp.8"|Project 98 file|
|"MSProject.mpp.9" |Project 2000–Project 2003 file|
|"MSProject.mpp.12"|Project 2007 file|
|"MSProject.odbc"|Open a project from an ODBC database|
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

Using the **FileOpenEx** method without specifying any arguments displays the **Open** dialog box with the list of enterprise projects if Project Professional is logged on Project Server. Using `FileOpenEx DoNotLoadFromEnterprise:=True` displays the **Open** dialog box for project files on the local computer.

If you use the **FileOpenEx** method to open a project that is published to Project Server, it opens the file from the Draft database. For example, to programmatically open a project named Project1 as read/write from Project Server, use the following command: `Application.FileOpenEx Name:="<>\Project1"`.

If you do not want to modify a project, set the _ReadOnly_ parameter to **True**. For example, to open Project2 as read-only, use the following command: `Application.FileOpenEx Name:="<>\Project2", ReadOnly:=True`. To save the file in the Draft database, use the  **Application.FileSave** method. To publish the file from the Draft to the Published database, so that changes are shown to other users, use the **Application.Publish** method.

The _Name_ parameter can contain a file name string or an ODBC data source name (DSN) and project name string. The syntax for a data source is <DataSourceName>\Projectname. The less than (<) and greater than (>) symbols must be included, and a backslash ( \ ) must separate the data source name from the project name. _DataSourceName_ itself can either be one of the ODBC data source names installed on the computer or a path and file name for a file-based database.


