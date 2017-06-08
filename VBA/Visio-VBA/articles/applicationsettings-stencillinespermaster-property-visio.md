---
title: ApplicationSettings.StencilLinesPerMaster Property (Visio)
keywords: vis_sdr.chm16251520
f1_keywords:
- vis_sdr.chm16251520
ms.prod: visio
api_name:
- Visio.ApplicationSettings.StencilLinesPerMaster
ms.assetid: 0d962d29-2cb5-5a9f-342f-1a35905a3438
ms.date: 06/08/2017
---


# ApplicationSettings.StencilLinesPerMaster Property (Visio)

For shapes on stencils in Microsoft Visio, determines how many lines of text of each shape's name can appear below the shape before the text is truncated and "..." is appended. Read/write.


## Syntax

 _expression_ . **StencilLinesPerMaster**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

 **Long**


## Remarks

Setting the  **StencilLinesPerMaster** property is equivalent to setting the **Lines per master** option under **Stencil spacing** on the **Advanced** tab in the **Visio Options** dialog box (click the **File** tab, and then click **Options**).

The minimum value for  **StencilLinesPerMaster** is 1 line per master and the maximum is 4. By default, Visio displays 2 lines per master.

This property affects the overall spacing of shapes on a stencil, which affects how many shapes users can see without scrolling.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **StencilCharactersPerLine** property to print the current number of stencil lines per master in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub StencilLinesPerMaster_Example() 
 
    Dim vsoApplicationSettings As Visio.ApplicationSettings 
    Dim lngLinesPerMaster As Long 
 
    Set vsoApplicationSettings = Visio.Application.Settings 
    lngLinesPerMaster = vsoApplicationSettings.StencilLinesPerMaster 
 
    Debug.Print lngLinesPerMaster 
 
End Sub
```


