---
title: ContentControl.DateDisplayFormat Property (Word)
keywords: vbawd10.chm266534925
f1_keywords:
- vbawd10.chm266534925
ms.prod: word
api_name:
- Word.ContentControl.DateDisplayFormat
ms.assetid: 11b2f24e-22d6-177c-4e2a-10c5ebefc477
ms.date: 06/08/2017
---


# ContentControl.DateDisplayFormat Property (Word)

Returns or sets a  **String** that represents the format in which dates are displayed. Read/write.


## Syntax

 _expression_ . **DateDisplayFormat**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

The default format is the format setting specified in Microsoft Word on the users' system, which usually depends on the location setting in Microsoft Windows. For example, the default format of dates for English (U.S.) is "mm/dd/yyyy". Use the  **DateDisplayFormat** property to specify a different date format.


## Example


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDate) 
 
objCC.Title = "Review Period End Date" 
objCC.DateDisplayFormat = "MMMM d, yyyy" 
objCC.DateStorageFormat = wdContentControlDateStorageDate 
objCC.Range.Text = "January 1, 2007"
```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

