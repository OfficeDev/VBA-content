---
title: ApplicationSettings.StencilCharactersPerLine Property (Visio)
keywords: vis_sdr.chm16251525
f1_keywords:
- vis_sdr.chm16251525
ms.prod: visio
api_name:
- Visio.ApplicationSettings.StencilCharactersPerLine
ms.assetid: e69c1c58-6383-f614-fcd4-d32505f53206
ms.date: 06/08/2017
---


# ApplicationSettings.StencilCharactersPerLine Property (Visio)

For shapes on stencils, determines approximately how many characters of each shape's name appear on each line before the text wraps to the next line. Read/write.


## Syntax

 _expression_ . **StencilCharactersPerLine**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **StencilCharactersPerLine** property is equivalent to setting the **Characters per line** option under **Stencil spacing** on the **Advanced** tab in the **Visio Options** dialog box (click the **File** tab, and then click **Options**).

The minimum value for  **StencilCharactersPerLine** is 5 characters per line and the maximum is 20. By default, Visio displays 12 characters per line.

This property affects the overall spacing of shapes on a stencil, which affects how many shapes the user can see without scrolling.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **StencilCharactersPerLine** property to print the current number of stencil characters per line in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub StencilCharactersPerLine_Example() 
 
    Dim vsoApplicationSettings As Visio.ApplicationSettings 
    Dim lngCharsPerLine As Long 
 
    Set vsoApplicationSettings = Visio.Application.Settings 
    lngCharsPerLine = vsoApplicationSettings.StencilCharactersPerLine 
 
    Debug.Print lngCharsPerLine 
 
End Sub
```


