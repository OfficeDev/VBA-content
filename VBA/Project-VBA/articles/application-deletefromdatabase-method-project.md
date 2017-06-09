---
title: Application.DeleteFromDatabase Method (Project)
keywords: vbapj.chm135
f1_keywords:
- vbapj.chm135
ms.prod: project-server
api_name:
- Project.Application.DeleteFromDatabase
ms.assetid: 22bed2ff-0e8b-e589-1479-06c482f296a9
ms.date: 06/08/2017
---


# Application.DeleteFromDatabase Method (Project)

Deletes a project stored in a database.


## Syntax

 _expression_. **DeleteFromDatabase**( ** _Name_**, ** _UserID_**, ** _DatabasePassWord_**, ** _FormatID_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the source file or data source to open, and the name of the project to delete from the database.|
| _UserID_|Optional|**String**| A user ID to use when accessing the database.|
| _DatabasePassWord_|Optional|**String**| A password to use when accessing the database.|
| _FormatID_|Optional|**String**|The file or database format. If Project recognizes the format of the file specified with Name, FormatID is ignored. Can be one of the following format strings:

|**Format String**|**Description**|
|:-----|:-----|
|"MSProject.mpd"|Project database|
|"MSProject.odbc"|ODBC database|
|"MSProject.mdb"|Microsoft Access database|
|

### Return Value

 **Boolean**


## Remarks

The Name argument must contain a file name string, or an ODBC data source name (DSN), and the project name string. The syntax for a data source is < _DataSourceName_ >\ _Projectname_. The less than (<) and greater than (>) symbols must be included, and a backslash ( \ ) must separate the data source name from the project name. The _DataSourceName_ itself can either be one of the ODBC data source names installed on the computer, a file DSN, or a path and file name for a file-based database.

In the following examples, _ [My Documents]_ is the full path of your My Documents folder, and _[Program Files]_ is the full path of your Program Files folder:

"<Corporate SQL Database>\Factory Construction" 

"<  _[My Documents]\_ PROJECT1.MDB>\System Roll-out Plan"

"<  _[Program Files]_ \Common Files\ODBC\Data Sources\Projects Database.dsn>\Project X"


## Example

The following example deletes projects from a Project database, as specified by the user.


```vb
Sub KillProjects() 
 Dim PathAndDB As String, ProjectName As String 
 Dim Continue As Long ' Used to store user response 
 
 Continue = vbYes ' Set to Yes so that loop runs 
 
 PathAndDB = InputBox$("Enter the path and file name of the Project" &; _ 
 " database to open, including extension: ") 
 
 Do Until Continue = vbNo 
 ProjectName = InputBox$("Enter the name of the project to delete: ") 
 DeleteFromDatabase "<" &; PathAndDB &; ">\" &; ProjectName, _ 
 FormatID:="MSProject.mpd" 
 Continue = MsgBox("Project " &; ProjectName &; " deleted from database." &; _ 
 vbCrLf &; vbCrLf &; "Delete another?", vbYesNo) 
 Loop 
 
End Sub
```


