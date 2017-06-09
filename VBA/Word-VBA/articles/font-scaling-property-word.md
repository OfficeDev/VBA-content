---
title: Font.Scaling Property (Word)
keywords: vbawd10.chm156369041
f1_keywords:
- vbawd10.chm156369041
ms.prod: word
api_name:
- Word.Font.Scaling
ms.assetid: 53f162cf-6de0-a142-50a5-fbdece3e7d16
ms.date: 06/08/2017
---


# Font.Scaling Property (Word)

Returns or sets the scaling percentage applied to the font. Read/write  **Long** .


## Syntax

 _expression_ . **Scaling**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

This property stretches or compresses text horizontally as a percentage of the current size (the scaling range is from 1 through 600).


## Example

This example horizontally stretches the text in the active document to 110 percent of its original size.


```vb
ActiveDocument.Content.Font.Scaling = 110
```

This example compresses the text in the first paragraph in Sales.doc to 90 percent of its original size.




```vb
With Documents("Sales.doc").Paragraphs(1).Range.Font 
 .Scaling = 90 
 .Bold = False 
End With
```


## See also


#### Concepts


[Font Object](font-object-word.md)

