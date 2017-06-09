---
title: DataTable.Application Property (Word)
keywords: vbawd10.chm46399636
f1_keywords:
- vbawd10.chm46399636
ms.prod: word
api_name:
- Word.DataTable.Application
ms.assetid: 4d966bc2-0548-8aeb-36f9-b71e6fbcf8f5
ms.date: 06/08/2017
---


# DataTable.Application Property (Word)

When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **[DataTable](datatable-object-word.md)** object.


## Example

The following example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveDocument 
If myObject.Application.Value = "Microsoft Word" Then 
 MsgBox "This is a Word Application object." 
Else 
 MsgBox "This is not a Word Application object." 
End If
```


## See also


#### Concepts


[DataTable Object](datatable-object-word.md)

