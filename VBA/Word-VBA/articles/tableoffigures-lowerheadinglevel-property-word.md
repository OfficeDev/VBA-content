---
title: TableOfFigures.LowerHeadingLevel Property (Word)
keywords: vbawd10.chm153157637
f1_keywords:
- vbawd10.chm153157637
ms.prod: word
api_name:
- Word.TableOfFigures.LowerHeadingLevel
ms.assetid: 5408cb26-a24b-4898-bf38-021357ce1633
ms.date: 06/08/2017
---


# TableOfFigures.LowerHeadingLevel Property (Word)

Returns or sets the ending heading level for a table of figures. Read/write  **Long** .


## Syntax

 _expression_ . **LowerHeadingLevel**

 _expression_ Required. A variable that represents a **[TableOfFigures](tableoffigures-object-word.md)** collection.


## Remarks

This property corresponds to the ending value used with the \o switch for a Table of Contents (TOC) field. Use the  **[UpperHeadingLevel](tableoffigures-upperheadinglevel-property-word.md)** property to set the starting heading level. For example, to set the TOC field syntax {TOC \o "1-3"}, set the **LowerHeadingLevel** property to 3 and the **UpperHeadingLevel** property to 1.


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

