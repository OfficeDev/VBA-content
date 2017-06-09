---
title: CurrentProject.Path Property (Access)
keywords: vbaac10.chm12718
f1_keywords:
- vbaac10.chm12718
ms.prod: access
api_name:
- Access.CurrentProject.Path
ms.assetid: 25f28502-b5fc-aafa-9189-eb091907a529
ms.date: 06/08/2017
---


# CurrentProject.Path Property (Access)

You can use the  **Path** property to determine the location where data is stored for a Microsoft Access project (.adp) or Microsoft Access database. Read-only **String**.


## Syntax

 _expression_. **Path**

 _expression_ A variable that represents a **CurrentProject** object.


## Remarks

You can use the  **Path** property to determine the location of information stored through the **[CurrentProject](currentproject-object-access.md)** or **[CodeProject](codeproject-object-access.md)** objects of a project or database.


## Example

The following example displays a message indicating the disk location of the current Access project or database.


```vb
MsgBox "The current database is located at " &; Application.CurrentProject.Path &; "." 
 

```


## See also


#### Concepts


[CurrentProject Object](currentproject-object-access.md)

