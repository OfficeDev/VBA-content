---
title: AllReports.Count Property (Access)
keywords: vbaac10.chm12681
f1_keywords:
- vbaac10.chm12681
ms.prod: access
api_name:
- Access.AllReports.Count
ms.assetid: e9c0908e-5eab-27d8-f301-c6d273555353
ms.date: 06/08/2017
---


# AllReports.Count Property (Access)

You can use the  **Count** property to determine the number of items in a specified collection. Read-only **Long**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents an **AllReports** object.


## Example

For example, if you want to determine the number of forms currently open or existing on the database, you would use the following code strings


```vb
' Determine the number of open forms. 
 
forms.count 
 
' Determine the number of forms (open or closed) 
' in the current database. 
 
currentproject.allforms.count
```


## See also


#### Concepts


[AllReports Collection](allreports-object-access.md)

