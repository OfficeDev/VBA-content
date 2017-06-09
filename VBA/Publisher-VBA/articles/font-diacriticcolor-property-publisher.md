---
title: Font.DiacriticColor Property (Publisher)
keywords: vbapb10.chm5374003
f1_keywords:
- vbapb10.chm5374003
ms.prod: publisher
api_name:
- Publisher.Font.DiacriticColor
ms.assetid: 6e9c816e-c7ae-c559-6b35-150a5abb820c
ms.date: 06/08/2017
---


# Font.DiacriticColor Property (Publisher)

Returns a  **[ColorFormat](colorformat-object-publisher.md)** object representing the 24-bit color used for diacritics in a right-to-left language publication.


## Syntax

 _expression_. **DiacriticColor**

 _expression_A variable that represents a  **Font** object.


### Return Value

ColorFormat


## Example

This example tests the text in the first story of the current publication to see if its color is red and it is formatted right-to-left.


```vb
Sub FontDiColor() 
 
 Dim fntDiColor As Font 
 
 Set fntDiColor = Application.ActiveDocument. _ 
 Stories(1).TextRange.Font 
 
 If fntDiColor.UseDiacriticColor = msoTrue And _ 
 fntDiColor.DiacriticColor.RGB = RGB(255, 0, 0) Then 
 MsgBox "Your text is red" 
 Else 
 MsgBox "This is not a right-to-left language" _ 
 &; " or your color is not red" 
 End If 
 
End Sub
```


