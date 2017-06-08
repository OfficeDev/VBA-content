---
title: Font.Shrink Method (Publisher)
keywords: vbapb10.chm5373991
f1_keywords:
- vbapb10.chm5373991
ms.prod: publisher
api_name:
- Publisher.Font.Shrink
ms.assetid: c5626ef2-5351-ab49-bf86-690587daed1f
ms.date: 06/08/2017
---


# Font.Shrink Method (Publisher)

Decreases the font size to the next available size. If the selection or range contains more than one font size, each size is decreased to the next available setting.


## Syntax

 _expression_. **Shrink**

 _expression_A variable that represents a  **Font** object.


## Remarks

Applying the  **Shrink** method to text that is already the smallest size allowed by Microsoft Publisher (0.5 point) has no effect.


## Example

This example inserts a line of increasingly smaller Z's in a new document.


```vb
Dim shpText As Shape 
Dim trTemp As TextRange 
Dim intCount As Integer 
 
Set shpText = ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=300, Height:=50) 
 
Set trTemp = shpText.TextFrame.TextRange 
 
With trTemp 
 .Font.Size = 45 
 .InsertAfter NewText:="ZZZZZZZZZZ" 
 For intCount = 2 To 10 
 .Characters(Start:=intCount, _ 
 Length:=11 - intCount).Font.Shrink 
 Next intCount 
End With
```


