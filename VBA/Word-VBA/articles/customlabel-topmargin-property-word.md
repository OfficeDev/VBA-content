---
title: CustomLabel.TopMargin Property (Word)
keywords: vbawd10.chm152371203
f1_keywords:
- vbawd10.chm152371203
ms.prod: word
api_name:
- Word.CustomLabel.TopMargin
ms.assetid: a1c783b1-08a9-ade0-6833-0b004a9f14ef
ms.date: 06/08/2017
---


# CustomLabel.TopMargin Property (Word)

Returns or sets the distance (in points) between the top edge of the page and the top boundary of the body text. Read/write  **Single** .


## Syntax

 _expression_ . **TopMargin**

 _expression_ Required. A variable that represents a **[CustomLabel](customlabel-object-word.md)** object.


## Example

This example creates a new custom label and sets several properties, including the top margin, and then it creates a new document using the custom labels.


```vb
Set newlbl = Application.MailingLabel. _ 
 CustomLabels.Add(Name:="My Label") 
With newlbl 
 .Height = InchesToPoints(1.25) 
 .NumberAcross = 2 
 .NumberDown = 7 
 .PageSize = wdCustomLabelLetter 
 .SideMargin = InchesToPoints(0) 
 .TopMargin = InchesToPoints(1) 
 .Width = InchesToPoints(4.25) 
End With 
Application.MailingLabel.CreateNewDocument Name:="My Label"
```


## See also


#### Concepts


[CustomLabel Object](customlabel-object-word.md)

