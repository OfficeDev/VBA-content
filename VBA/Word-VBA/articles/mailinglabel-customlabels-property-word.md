---
title: MailingLabel.CustomLabels Property (Word)
keywords: vbawd10.chm152502280
f1_keywords:
- vbawd10.chm152502280
ms.prod: word
api_name:
- Word.MailingLabel.CustomLabels
ms.assetid: c4bad9e7-8da9-d469-4d49-a3b43c5cc4de
ms.date: 06/08/2017
---


# MailingLabel.CustomLabels Property (Word)

Returns a  **[CustomLabels](customlabels-object-word.md)** collection that represents the available custom mailing labels. Read-only.


## Syntax

 _expression_ . **CustomLabels**

 _expression_ A variable that represents a **[MailingLabel](mailinglabel-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example creates a new custom label named "AdminAddress" and then creates a page of mailing labels using a predefined return address.


```vb
Dim strAddress As String 
Dim labelNew As CustomLabel 
 
strAddress = "Administration" &; vbCr &; "Mail Stop 22-16" 
 
Set labelNew = Application.MailingLabel _ 
 .CustomLabels.Add(Name:="AdminAddress", DotMatrix:= False) 
 
With labelNew 
 .Height = InchesToPoints(0.5) 
 .Width = InchesToPoints(1) 
 .HorizontalPitch = InchesToPoints(2.06) 
 .VerticalPitch = InchesToPoints(0.5) 
 .NumberAcross = 4 
 .NumberDown = 20 
 .PageSize = wdCustomLabelLetter 
 .SideMargin = InchesToPoints(0.28) 
 .TopMargin = InchesToPoints(0.5) 
End With 
 
Application.MailingLabel.CreateNewDocument _ 
 Name:="AdminAddress", Address:=strAddress
```


## See also


#### Concepts


[MailingLabel Object](mailinglabel-object-word.md)

