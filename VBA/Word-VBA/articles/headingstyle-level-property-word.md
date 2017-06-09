---
title: HeadingStyle.Level Property (Word)
keywords: vbawd10.chm160104450
f1_keywords:
- vbawd10.chm160104450
ms.prod: word
api_name:
- Word.HeadingStyle.Level
ms.assetid: 6a322568-ebcb-4ee4-2123-a94b3f97bc1a
ms.date: 06/08/2017
---


# HeadingStyle.Level Property (Word)

Returns or sets the level for the heading style in a table of contents or table of figures. Read/write  **Integer** .


## Syntax

 _expression_ . **Level**

 _expression_ Required. A variable that represents a **[HeadingStyle](headingstyle-object-word.md)** object.


## Example

This example adds a table of contents at the insertion point in the active document, and then it changes the levels for the heading styles.


```vb
ActiveDocument.TablesOfContents.Add _ 
 Range:=Selection.Range, _ 
 RightAlignPageNumbers:=True, _ 
 UseHeadingStyles:=True, _ 
 UpperHeadingLevel:=1, _ 
 LowerHeadingLevel:=3, _ 
 IncludePageNumbers:=True, _ 
 TableID:=wdTOCFormal 
With ActiveDocument.TablesOfContents(1).HeadingStyles 
 .Add Style:="Title", Level:=1 
 .Add Style:="SubTitle", Level:=2 
 .Add Style:="List Bullet", Level:=3 
End With 
With ActiveDocument.TablesOfContents(1) 
 .HeadingStyles(1).Level = 2 
 .HeadingStyles(2).Level = 4 
 .HeadingStyles(3).Level = 6 
End With
```


## See also


#### Concepts


[HeadingStyle Object](headingstyle-object-word.md)

