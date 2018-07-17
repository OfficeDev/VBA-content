---
title: CodeProject.BaseConnectionString Property (Access)
keywords: vbaac10.chm12713
f1_keywords:
- vbaac10.chm12713
ms.prod: access
api_name:
- Access.CodeProject.BaseConnectionString
ms.assetid: 118da929-8e30-25eb-c940-27ce924a88f7
ms.date: 06/08/2017
---


# CodeProject.BaseConnectionString Property (Access)

You can use the  **BaseConnectionString** property to return the base connection string for the specified object. Read-only **String**.


## Syntax

 _expression_. **BaseConnectionString**

 _expression_ A variable that represents a **CodeProject** object.


## Remarks

The  **BaseConnectionString** property returns the connection string that was set through the **OpenConnection** method or by clicking **Connection** on the **File** menu. When making a connection, Microsoft Access project modifies the **BaseConnectionString** property for use with the ADO environment.


## Example

The following example displays the  **BaseConnectionString** property setting of the current project:


```vb
Public Sub ShowConnectString() 
 
 Dim objCurrent As Object 
 
 Set objCurrent = Application.CurrentProject 
 MsgBox "The current base connection is " &; objCurrent.BaseConnectionString 
 
End Sub
```


## See also


#### Concepts


[CodeProject Object](codeproject-object-access.md)

