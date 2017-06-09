---
title: CellBorder.Color Property (Publisher)
keywords: vbapb10.chm5242882
f1_keywords:
- vbapb10.chm5242882
ms.prod: publisher
api_name:
- Publisher.CellBorder.Color
ms.assetid: 59a43522-f0df-fe1a-6e35-19cb012b103f
ms.date: 06/08/2017
---


# CellBorder.Color Property (Publisher)

Returns a  **[ColorFormat](colorformat-object-publisher.md)** object representing the color information for the specified object.


## Syntax

 _expression_. **Color**

 _expression_A variable that represents a  **CellBorder** object.


## Example

This example tests the font color of the first story in the active document and tells the user if the font color is black or not.


```vb
Sub FontColor() 
 
 If Application.ActiveDocument.Stories(1) _ 
 .TextRange.Font.Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) Then 
 MsgBox "Your font color is black" 
 Else 
 MsgBox "Your font color is not black" 
 End If 
 
End Sub
```


