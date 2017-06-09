---
title: TextEffectFormat.FontBold Property (Excel)
keywords: vbaxl10.chm118003
f1_keywords:
- vbaxl10.chm118003
ms.prod: excel
api_name:
- Excel.TextEffectFormat.FontBold
ms.assetid: 19773cce-32d3-b07f-4650-5a19a4aa469a
ms.date: 06/08/2017
---


# TextEffectFormat.FontBold Property (Excel)

 **True** if the font in the specified WordArt is bold. Read/write **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **FontBold**

 _expression_ A variable that represents a **TextEffectFormat** object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse**|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** The specified WordArt is bold.|

## Example

This example sets the font to bold for shape three on  `myDocument` if the shape is WordArt.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
    If .Type = msoTextEffect Then 
        .TextEffect.FontBold = msoTrue 
    End If 
End With
```


## See also


#### Concepts


[TextEffectFormat Object](texteffectformat-object-excel.md)

