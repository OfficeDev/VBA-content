---
title: TextRange.MajorityFont Property (Publisher)
keywords: vbapb10.chm5308467
f1_keywords:
- vbapb10.chm5308467
ms.prod: publisher
api_name:
- Publisher.TextRange.MajorityFont
ms.assetid: b0007ebc-ed0b-aab8-49fe-76353efbc1d2
ms.date: 06/08/2017
---


# TextRange.MajorityFont Property (Publisher)

Returns a  **[Font](font-object-publisher.md)** object that represents the font name most in use in a text range.


## Syntax

 _expression_. **MajorityFont**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

Font


## Example

This example creates a new text box, fills it with text, checks if the font most in use is Tahoma, and if it isn't, changes the font to Tahoma.


```vb
Sub SetFontName() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 If .MajorityFont <> "Tahoma" Then _ 
 .Font.Name = "Tahoma" 
 End With 
End Sub
```


