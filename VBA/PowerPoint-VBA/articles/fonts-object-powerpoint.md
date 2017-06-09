---
title: Fonts Object (PowerPoint)
keywords: vbapp10.chm528000
f1_keywords:
- vbapp10.chm528000
ms.prod: powerpoint
api_name:
- PowerPoint.Fonts
ms.assetid: 1a8f44ea-515f-5eb9-eab5-6204d9b1d5bc
ms.date: 06/08/2017
---


# Fonts Object (PowerPoint)

A collection of all the  **[Font](font-object-powerpoint.md)** objects in the specified presentation.


## Remarks

Each  **Font** object represents a font that's used in the presentation.


## Example

Use the [Fonts](presentation-fonts-property-powerpoint.md) property to return the **Fonts** collection. The following example displays the number of fonts used in the active presentation.


```vb
MsgBox ActivePresentation.Fonts.Count
```

Use  **Fonts** (index), where index is the font's name or index number, to return a single **Font** object. The following example checks to see whether font one in the active presentation is embedded in the presentation.




```
If ActivePresentation.Fonts(1).Embedded = True Then 
    MsgBox "Font 1 is embedded"
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

