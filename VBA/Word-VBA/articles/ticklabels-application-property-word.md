---
title: TickLabels.Application Property (Word)
keywords: vbawd10.chm167051412
f1_keywords:
- vbawd10.chm167051412
ms.prod: word
api_name:
- Word.TickLabels.Application
ms.assetid: c6265ab0-d489-1c78-3e1d-9fc5affe5e1c
ms.date: 06/08/2017
---


# TickLabels.Application Property (Word)

When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **[TickLabels](ticklabels-object-word.md)** object.


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


[TickLabels Object](ticklabels-object-word.md)

