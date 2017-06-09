---
title: TableOfFigures.UpperHeadingLevel Property (Word)
keywords: vbawd10.chm153157638
f1_keywords:
- vbawd10.chm153157638
ms.prod: word
api_name:
- Word.TableOfFigures.UpperHeadingLevel
ms.assetid: bfda7885-8aec-96d7-2bdf-93ddd2804385
ms.date: 06/08/2017
---


# TableOfFigures.UpperHeadingLevel Property (Word)

Returns or sets the starting heading level for a table of figures. Read/write  **Long** .


## Syntax

 _expression_ . **UpperHeadingLevel**

 _expression_ Required. A variable that represents a **[TableOfFigures](tableoffigures-object-word.md)** collection.


## Remarks

This property corresponds to the starting value used with the \o switch for a Table of Contents (TOC) field. Use the  **[LowerHeadingLevel](tableoffigures-lowerheadinglevel-property-word.md)** property to set the ending heading level. For example, to set the TOC field syntax {TOC \o "1-3"}, set the **LowerHeadingLevel** property to 3 and the **UpperHeadingLevel** property to 1.


## Example

This example formats the first table of table of figures in the active document to compile all headings that are formatted with either the Heading 8 or Heading 9 style.


```vb
If ActiveDocument.TablesOfFigures.Count >= 1 Then 
 With ActiveDocument.TablesOfFigures(1) 
 .UseHeadingStyles = True 
 .UseFields = False 
 .UpperHeadingLevel = 8 
 .LowerHeadingLevel = 9 
 End With 
End If
```


## See also


#### Concepts


[TableOfFigures Object](tableoffigures-object-word.md)

