---
title: ApplicationSettings.FreeformDrawingSmoothing Property (Visio)
keywords: vis_sdr.chm16251785
f1_keywords:
- vis_sdr.chm16251785
ms.prod: visio
api_name:
- Visio.ApplicationSettings.FreeformDrawingSmoothing
ms.assetid: 55526b81-324a-8c6f-1654-bf7e1244ccf2
ms.date: 06/08/2017
---


# ApplicationSettings.FreeformDrawingSmoothing Property (Visio)

Determines how precisely mouse movements are smoothed when drawing a spline. Read/write.


## Syntax

 _expression_ . **FreeformDrawingSmoothing**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **FreeformDrawingSmoothing** property is equivalent to setting the **Smoothing** option on the **Advanced** tab in the **Visio Options** dialog box (click the **File** tab, and then click **Options**).

Possible values for the  **FreeformDrawingSmoothing** property range from 0 ( **Tight**) to 10 ( **Loose**). The default is 5.


