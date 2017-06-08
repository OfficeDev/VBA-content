---
title: PageSetup.TopMargin Property (Word)
keywords: vbawd10.chm158400612
f1_keywords:
- vbawd10.chm158400612
ms.prod: word
api_name:
- Word.PageSetup.TopMargin
ms.assetid: c7c8d859-e82b-5170-eadb-95a6e5895f83
ms.date: 06/08/2017
---


# PageSetup.TopMargin Property (Word)

Returns or sets the distance (in points) between the top edge of the page and the top boundary of the body text. Read/write  **Single** .


## Syntax

 _expression_ . **TopMargin**

 _expression_ Required. A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example sets the top margin to 72 points (1 inch) for the first section in the active document.


```vb
ActiveDocument.Sections(1).PageSetup.TopMargin = 72
```

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


[PageSetup Object](pagesetup-object-word.md)

