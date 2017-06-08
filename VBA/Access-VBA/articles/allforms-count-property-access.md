---
title: AllForms.Count Property (Access)
keywords: vbaac10.chm12681
f1_keywords:
- vbaac10.chm12681
ms.prod: access
api_name:
- Access.AllForms.Count
ms.assetid: 1540145e-541d-10fc-249b-9fadc6861a11
ms.date: 06/08/2017
---


# AllForms.Count Property (Access)

You can use the  **Count** property to determine the number of items in a specified collection. Read-only **Long**.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents an **AllForms** object.


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


[AllForms Collection](allforms-object-access.md)

