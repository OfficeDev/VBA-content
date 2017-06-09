---
title: ColorFormat.Brightness Property (Excel)
ms.prod: excel
api_name:
- Excel.ColorFormat.Brightness
ms.assetid: 36428885-90c0-327f-2ecc-5160ae6263cd
ms.date: 06/08/2017
---


# ColorFormat.Brightness Property (Excel)

Returns or sets the luminosity of the specified object. Read/write


## Syntax

 _expression_ . **Brightness**

 _expression_ A variable that represents a **[ColorFormat](colorformat-object-excel.md)** object.


### Return Value

 **Single**


## Remarks

 The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest).


## Example

The following code example sets the brightness of the fill color for the first shape in the active worksheet.


```vb
ActiveSheet.Shapes(1).Fill.ForeColor.Brightness = 0.5
```


## See also


#### Concepts


[ColorFormat Object](colorformat-object-excel.md)

