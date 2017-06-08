---
title: TextEffectFormat Object (Word)
keywords: vbawd10.chm2511
f1_keywords:
- vbawd10.chm2511
ms.prod: word
api_name:
- Word.TextEffectFormat
ms.assetid: b274e5be-ed5b-7d63-aa4b-1d67b63e7c0b
ms.date: 06/08/2017
---


# TextEffectFormat Object (Word)

Contains properties and methods that apply to WordArt objects.


## Remarks

Use the  **TextEffect** property to return a **TextEffectFormat** object. The following example sets the font name and formatting for shape one on the active document. For this example to work, shape one must be a WordArt object.


```vb
With ActiveDocument.Shapes(1).TextEffect 
 .FontName = "Courier New" 
 .FontBold = True 
 .FontItalic = True 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


