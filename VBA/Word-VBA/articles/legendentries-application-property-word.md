---
title: LegendEntries.Application Property (Word)
keywords: vbawd10.chm6815892
f1_keywords:
- vbawd10.chm6815892
ms.prod: word
api_name:
- Word.LegendEntries.Application
ms.assetid: 41c01d34-0f89-c898-4b8a-43daf05d9a8d
ms.date: 06/08/2017
---


# LegendEntries.Application Property (Word)

When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **[LegendEntries](legendentries-object-word.md)** object.


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


[LegendEntries Object](legendentries-object-word.md)

