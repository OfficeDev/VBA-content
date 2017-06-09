---
title: ContentControl.DateStorageFormat Property (Word)
keywords: vbawd10.chm266534932
f1_keywords:
- vbawd10.chm266534932
ms.prod: word
api_name:
- Word.ContentControl.DateStorageFormat
ms.assetid: c69d3f01-725e-8b64-147b-ca8a146b7419
ms.date: 06/08/2017
---


# ContentControl.DateStorageFormat Property (Word)

Returns or sets a  **[WdContentControlDateStorageFormat](wdcontentcontroldatestorageformat-enumeration-word.md)** that represents the format for storage and retrieval of dates when a date content control is bound to the XML data store of the active document. Read/write.


## Syntax

 _expression_ . **DateStorageFormat**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

The  **DateStorageFormat** property allows you to store dates in date format, date/time format, or text format.


## Example

The following example adds a date content control to the active document and specifies the date, the date display format, and the date storage format.


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

