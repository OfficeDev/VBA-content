---
title: TextEffectFormat Object (Excel)
keywords: vbaxl10.chm118000
f1_keywords:
- vbaxl10.chm118000
ms.prod: excel
api_name:
- Excel.TextEffectFormat
ms.assetid: 7fe03721-6a45-569e-add4-fc8849c99535
ms.date: 06/08/2017
---


# TextEffectFormat Object (Excel)

Contains properties and methods that apply to WordArt objects.


## Remarks

Use the  **[TextEffect](shape-texteffect-property-excel.md)** property to return a **TextEffectFormat** object.


## Example

 The following example sets the font name and formatting for shape one on _myDocument_ . For this example to work, shape one must be a WordArt object.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).TextEffect 
 .FontName = "Courier New" 
 .FontBold = True 
 .FontItalic = True 
End With 

```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


