---
title: Range.IsEndOfRowMark Property (Word)
keywords: vbawd10.chm157155635
f1_keywords:
- vbawd10.chm157155635
ms.prod: word
api_name:
- Word.Range.IsEndOfRowMark
ms.assetid: 0b1a7638-75ea-fb03-3a52-8bc759794408
ms.date: 06/08/2017
---


# Range.IsEndOfRowMark Property (Word)

 **True** if the specified range is collapsed and is located at the end-of-row mark in a table. Read-only **Boolean** .


## Syntax

 _expression_ . **IsEndOfRowMark**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

This property is the equivalent of the following expression:


```vb
ActiveDocument.Range.Information(wdAtEndOfRowMarker)
```


## Example

This example collapses the selection and selects the current row if the insertion point is at the end of the row (just before the end-of-row mark).


```vb
ActiveDocument.Range.Collapse Direction:=wdCollapseEnd 
If ActiveDocument.Range.IsEndOfRowMark = True Then 
 ActiveDocument.Range.Rows(1).Select 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

