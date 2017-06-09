---
title: Point.Application Property (Word)
keywords: vbawd10.chm262144148
f1_keywords:
- vbawd10.chm262144148
ms.prod: word
api_name:
- Word.Point.Application
ms.assetid: 9aa40408-6ec0-ca1f-4cb8-8de01729255a
ms.date: 06/08/2017
---


# Point.Application Property (Word)

When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **[Point](point-object-word.md)** object.


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


[Point Object](point-object-word.md)

