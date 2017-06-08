---
title: TableOfFigures.UseHeadingStyles Property (Word)
keywords: vbawd10.chm153157636
f1_keywords:
- vbawd10.chm153157636
ms.prod: word
api_name:
- Word.TableOfFigures.UseHeadingStyles
ms.assetid: f4688cfb-df6e-5567-9652-54c25fd4e410
ms.date: 06/08/2017
---


# TableOfFigures.UseHeadingStyles Property (Word)

 **True** if built-in heading styles are used to create a table of figures. Read/write **Boolean** .


## Syntax

 _expression_ . **UseHeadingStyles**

 _expression_ Required. A variable that represents a **[TableOfFigures](tableoffigures-object-word.md)** collection.


## Example

This example adds a table of figures in place of the selection and then formats the table to compile entries from TC fields.


```vb
With ActiveDocument.TablesOfFigures.Add(Range:=Selection.Range) 
 .UseHeadingStyles = False 
 .UseFields = True 
End With
```


## See also


#### Concepts


[TableOfFigures Object](tableoffigures-object-word.md)

