---
title: Font.SizeBi Property (Publisher)
keywords: vbapb10.chm5373958
f1_keywords:
- vbapb10.chm5373958
ms.prod: publisher
api_name:
- Publisher.Font.SizeBi
ms.assetid: 1e9100e7-efa4-a7aa-69af-39c550a0b046
ms.date: 06/08/2017
---


# Font.SizeBi Property (Publisher)

Returns or sets a  **Variant** value representing the size, in points, of the **Font** object for text in a right-to-left language. Valid range is 0.5 points to 999.5 points. Read/write.


## Syntax

 _expression_. **SizeBi**

 _expression_A variable that represents a  **Font** object.


### Return Value

Variant


## Example

This example tests the text in the second story. If it is in a right-to-left language, larger than 12 point, and italic, the text is set to bold.


```vb
Sub SizeBiIfBig() 
 
 Dim fntSize As Font 
 
 Set fntSize = Application.ActiveDocument.Stories(2).TextRange.Font 
 With fntSize 
 If .SizeBi > 12 And .ItalicBi = msoTrue Then 
 .BoldBi = msoTrue 
 Else 
 MsgBox "The font size is 12 points or less " &; _ 
 ", not bold, or this is not in a right-to-left language." 
 End If 
 End With 
 
End Sub
```


