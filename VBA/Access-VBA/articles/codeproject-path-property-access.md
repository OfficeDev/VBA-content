---
title: CodeProject.Path Property (Access)
keywords: vbaac10.chm12718
f1_keywords:
- vbaac10.chm12718
ms.prod: access
api_name:
- Access.CodeProject.Path
ms.assetid: 3d811cc3-ebb3-3cbc-fc3d-e1ab40ceea27
ms.date: 06/08/2017
---


# CodeProject.Path Property (Access)

You can use the  **Path** property to determine the location where data is stored for a Microsoft Access project (.adp) or Microsoft Access database. Read-only **String**.


## Syntax

 _expression_. **Path**

 _expression_ A variable that represents a **CodeProject** object.


## Remarks

You can use the  **Path** property to determine the location of information stored through the **[CurrentProject](currentproject-object-access.md)** or **[CodeProject](codeproject-object-access.md)** objects of a project or database.


## Example

The following example displays a message indicating the disk location of the current Access project or database.


```vb
MsgBox "The current database is located at " &; Application.CurrentProject.Path &; "." 
 

```


## See also


#### Concepts


[CodeProject Object](codeproject-object-access.md)

