---
title: Application.FileFormatID Property (Project)
keywords: vbapj.chm131218
f1_keywords:
- vbapj.chm131218
ms.prod: project-server
api_name:
- Project.Application.FileFormatID
ms.assetid: 86a6a5ce-6508-f1ad-b9cc-fb86fd96e410
ms.date: 06/08/2017
---


# Application.FileFormatID Property (Project)

Gets a value that indicates the file format for the specified project. Possible formats are those that Project can directly open as a project file. Read-only  **String**.


## Syntax

 _expression_. **FileFormatID**( ** _Name_**, ** _UserID_**, ** _DatabasePassWord_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|Name of the project file.|
| _UserID_|Optional|**Variant**|User name or identification number for project files stored in an ODBC database.|
| _DatabasePassWord_|Optional|**Variant**|Password for the ODBC database.|

## Remarks

The specified file must be a project file in the current directory. The  **FileFormatID** property in Project can be one of the following strings, for the specified file format:


- MSProject.MPP.14, for a standard Project or Project file
    
- MSProject.MPP.12, for a standard Office Project 2007 file
    
- MSProject.MPP.9, for a standard Project 2000, Project 2002, or Office Project 2003 file
    
- MSProject.MPT.14, for a Project or Project template
    
- MSProject.MPT.12, for an Office Project 2007 template
    
- MSProject.ACE.14, for a project saved as an Office Excel 2007 or Excel .xlsx file
    
- MSProject.ACEB.14, for a project saved as an Office Excel 2007 or Excel .xlsb file
    
- MSProject.XLS5.9, for a project saved as an .xls file in Excel 97 to Excel 2003
    
- MSProject.ODBC.9, for a project stored in an ODBC database, such as an Access database (.mdb file)
    



 **Note**  For backward compatibility with project files that are accessible only through ODBC (Open Database Connectivity), Project can open files using an ODBC connection. To save any changes after you open the file, however, you must save the file in another format on the local computer or to Project Server.


