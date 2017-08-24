---
title: Plate.Color Property (Publisher)
keywords: vbapb10.chm2883587
f1_keywords:
- vbapb10.chm2883587
ms.prod: publisher
api_name:
- Publisher.Plate.Color
ms.assetid: 4c4897f5-90bb-cb84-e9b8-47df1a912916
ms.date: 06/08/2017
---


# Plate.Color Property (Publisher)

Returns a  **[ColorFormat](colorformat-object-publisher.md)** object representing the color information for the specified object.


## Syntax

 _expression_. **Color**

 _expression_A variable that represents a  **Plate** object.


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


