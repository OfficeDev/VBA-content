---
title: TabStop.Position Property (Word)
keywords: vbawd10.chm156500070
f1_keywords:
- vbawd10.chm156500070
ms.prod: word
api_name:
- Word.TabStop.Position
ms.assetid: f44ce39b-34e6-992b-fe50-be53bd6f53bf
ms.date: 06/08/2017
---


# TabStop.Position Property (Word)

Returns or sets the position of a tab stop relative to the left margin. Read/write  **Single** .


## Syntax

 _expression_ . **Position**

 _expression_ Required. A variable that represents a **[TabStop](tabstop-object-word.md)** object.


## Example

This example adds a right tab stop to the selected paragraphs 2 inches from the left margin. The position of the tab stop is then displayed in a message box.


```vb
With Selection.Paragraphs.TabStops 
 .ClearAll 
 .Add Position:=InchesToPoints(2), Alignment:=wdAlignTabRight 
 MsgBox .Item(1).Position &; " or " &; _ 
 PointsToInches(.Item(1).Position) &; " inches" 
End With
```


## See also


#### Concepts


[TabStop Object](tabstop-object-word.md)

