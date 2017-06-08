---
title: DataLabels.Application Property (Word)
keywords: vbawd10.chm207487124
f1_keywords:
- vbawd10.chm207487124
ms.prod: word
api_name:
- Word.DataLabels.Application
ms.assetid: c0a70d81-7a14-5e68-b257-1784699b3571
ms.date: 06/08/2017
---


# DataLabels.Application Property (Word)

When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-word.md)** object.


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


[DataLabels Object](datalabels-object-word.md)

