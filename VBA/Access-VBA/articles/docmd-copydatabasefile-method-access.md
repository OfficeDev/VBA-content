---
title: DoCmd.CopyDatabaseFile Method (Access)
keywords: vbaac10.chm5088
f1_keywords:
- vbaac10.chm5088
ms.prod: access
api_name:
- Access.DoCmd.CopyDatabaseFile
ms.assetid: 15a820d9-fbcb-d803-d58a-5718924e6c73
ms.date: 06/08/2017
---


# DoCmd.CopyDatabaseFile Method (Access)

Copies the database connected to the current project to a Microsoft SQL Server database file for export.


## Syntax

 _expression_. **CopyDatabaseFile**( ** _DatabaseFileName_**, ** _OverwriteExistingFile_**, ** _DisconnectAllUsers_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DatabaseFileName_|Required|**Variant**|The name of the file (and path) to which the current database is copied. If no path is specified, the current directory is used.|
| _OverwriteExistingFile_|Optional|**Variant**|Determines whether Microsoft Access overwrites the file specified by  _DatabaseFileName_.  **True** to overwrite the existing file. If the file doesn't already exist, this argument is ignored|
| _DisconnectAllUsers_|Optional|**Variant**|Determines whether Access disconnects any users connected to the current database in order to make the copy.  **True** to disconnect other users before copying the database file.|

## Remarks

The file name of the copy must have an .mdf extension in order to be recognized as a SQL Server database file.

The method fails and an error occurs if any of the following occurs:


-  _DisconnectAllUsers_ is **True** but Access is unable to log off other users.
    
- The method cancels a save operation by any open design sessions.
    
- The destination file exists but  _OverwriteExistingFile_ was not set to **True**.
    
- The destination file exists, but is in use by another application.
    
- Access could not reconnect the original .mdf file.
    
- The current user for the Access project doesn?t have system administrator privileges for the database server.
    

## Example

This example copies the database connected to the current project to a SQL Server database file. If the file exists already, Access overwrites it, and any other users connected to the database are disconnected before the copy is made.


```vb
DoCmd.CopySQLDatabaseFile _ 
 DatabaseFileName:="C:\Export\Sales.mdf", _ 
 OverwriteExistingFile:=True, _ 
 DisconnectAllUsers:=True
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

