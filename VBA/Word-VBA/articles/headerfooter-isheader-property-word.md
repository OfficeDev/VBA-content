---
title: HeaderFooter.IsHeader Property (Word)
keywords: vbawd10.chm159711235
f1_keywords:
- vbawd10.chm159711235
ms.prod: word
api_name:
- Word.HeaderFooter.IsHeader
ms.assetid: 66c098ed-d0d6-cf58-e26a-b031bc7a6cab
ms.date: 06/08/2017
---


# HeaderFooter.IsHeader Property (Word)

 **True** if the specified **HeaderFooter** object is a header. Read-only **Boolean** .


## Syntax

 _expression_ . **IsHeader**

 _expression_ An expression that returns a **[HeaderFooter](headerfooter-object-word.md)** object.


## Example

This example selects the footer and adds a page number.


```vb
With ActiveDocument.ActiveWindow.ActivePane.View 
 .Type = wdPrintView 
 .SeekView = wdSeekCurrentPageHeader 
End With 
 
If Selection.HeaderFooter.IsHeader = True Then 
 ActiveDocument.ActiveWindow.ActivePane.View _ 
 .SeekView = wdSeekCurrentPageFooter 
End If 
 
Selection.HeaderFooter.PageNumbers.Add
```


## See also


#### Concepts


[HeaderFooter Object](headerfooter-object-word.md)

