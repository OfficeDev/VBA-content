---
title: CodeProject.FileFormat Property (Access)
keywords: vbaac10.chm12725
f1_keywords:
- vbaac10.chm12725
ms.prod: access
api_name:
- Access.CodeProject.FileFormat
ms.assetid: f72e13d7-58dc-3710-b1e7-798a71601ceb
ms.date: 06/08/2017
---


# CodeProject.FileFormat Property (Access)

Returns an  **[AcFileFormat](acfileformat-enumeration-access.md)** constant indicating the Microsoft Access version format of the specified project. Read-only.


## Syntax

 _expression_. **FileFormat**

 _expression_ A variable that represents a **CodeProject** object.


## Remarks

Use the  **ConvertAccessProject** method to convert an Access project from one version to another.


## Example

This example evaluates the current project and displays a message indicating to which version of Microsoft Access its file format corresponds.


```vb
Dim strFormat As String 
 
Select Case CurrentProject.FileFormat 
 Case acFileFormatAccess2 
 strFormat = "Microsoft Access 2" 
 Case acFileFormatAccess95 
 strFormat = "Microsoft Access 95" 
 Case acFileFormatAccess97 
 strFormat = "Microsoft Access 97" 
 Case acFileFormatAccess2000 
 strFormat = "Microsoft Access 2000" 
 Case acFileFormatAccess2002 
 strFormat = "Access 2002 - 2003" 
 Case acFileFormatAccess12 
 strFormat = "Microsoft Access 2007" 
End Select 
 
MsgBox "This is a " &; strFormat &; " project."
```


## See also


#### Concepts


[CodeProject Object](codeproject-object-access.md)

