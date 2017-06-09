---
title: TableOfFigures.TableID Property (Word)
keywords: vbawd10.chm153157642
f1_keywords:
- vbawd10.chm153157642
ms.prod: word
api_name:
- Word.TableOfFigures.TableID
ms.assetid: b7154038-2af5-2542-e1d8-c4002ec96cdf
ms.date: 06/08/2017
---


# TableOfFigures.TableID Property (Word)

Returns or sets a one-letter identifier that is used to build a table of figures from TOC fields. Read/write  **String** .


## Syntax

 _expression_ . **TableID**

 _expression_ Required. A variable that represents a **[TableOfFigures](tableoffigures-object-word.md)** collection.


## Remarks

This property corresponds to the \f switch for a TOC field. For example, "T" builds a table of figures from TOC fields using the table identifier T.


## Example

This example adds a table of figures at the beginning of the active document and then formats the table to compile "B" entries.


```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
Set myTOF = ActiveDocument.TablesOfFigures.Add(Range:=myRange) 
With myTOF 
 .UseFields = True 
 .UseHeadingStyles = False 
 .TableID = "B" 
 .Caption = "" 
End With
```


## See also


#### Concepts


[TableOfFigures Object](tableoffigures-object-word.md)

