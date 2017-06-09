---
title: Rows.Alignment Property (Word)
keywords: vbawd10.chm155975684
f1_keywords:
- vbawd10.chm155975684
ms.prod: word
api_name:
- Word.Rows.Alignment
ms.assetid: 0a3352eb-6618-1721-6261-11adad48707c
ms.date: 06/08/2017
---


# Rows.Alignment Property (Word)

Returns or sets a  **WdRowAlignment** constant that represents the alignment for the specified rows. Read/write.


## Syntax

 _expression_ . **Alignment**

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


## Example

This example centers all the rows in the first table of the active document.


```vb
Sub CenterRows() 
 ActiveDocument.Tables(1).Rows _ 
 .Alignment = wdAlignRowCenter 
End Sub
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

