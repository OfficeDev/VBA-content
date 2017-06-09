---
title: CustomLabel.PageSize Property (Word)
keywords: vbawd10.chm152371212
f1_keywords:
- vbawd10.chm152371212
ms.prod: word
api_name:
- Word.CustomLabel.PageSize
ms.assetid: b2a9e63e-041a-d4fc-6135-0e1e294886a2
ms.date: 06/08/2017
---


# CustomLabel.PageSize Property (Word)

Returns or sets the page size for the specified custom mailing label. Read/write  **WdCustomLabelPageSize** .


## Syntax

 _expression_ . **PageSize**

 _expression_ Required. A variable that represents a **[CustomLabel](customlabel-object-word.md)** object.


## Remarks

Some of the  **WdCustomLabelPageSize** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## Example

This example creates a new custom label named "Home Address" and then sets various properties for the label, including the page size.


```vb
Set myLabel = Application.MailingLabel _ 
 .CustomLabels.Add(Name:="Home Address", DotMatrix:=False) 
With myLabel 
 .Height = InchesToPoints(0.5) 
 .HorizontalPitch = InchesToPoints(2.06) 
 .NumberAcross = 4 
 .NumberDown = 20 
 .PageSize = wdCustomLabelLetter 
 .SideMargin = InchesToPoints(0.28) 
 .TopMargin = InchesToPoints(0.5) 
 .VerticalPitch = InchesToPoints(0.5) 
 .Width = InchesToPoints(1.75) 
End With
```


## See also


#### Concepts


[CustomLabel Object](customlabel-object-word.md)

