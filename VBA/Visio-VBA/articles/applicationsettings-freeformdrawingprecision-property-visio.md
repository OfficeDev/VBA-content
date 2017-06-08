---
title: ApplicationSettings.FreeformDrawingPrecision Property (Visio)
keywords: vis_sdr.chm16251790
f1_keywords:
- vis_sdr.chm16251790
ms.prod: visio
api_name:
- Visio.ApplicationSettings.FreeformDrawingPrecision
ms.assetid: 3822238b-cd63-1883-88a6-894b289765d7
ms.date: 06/08/2017
---


# ApplicationSettings.FreeformDrawingPrecision Property (Visio)

Determines the margin of error allowed when the  **Freeform** tool is drawing a straight line before it switches to drawing a spline. Read/write.


## Syntax

 _expression_ . **FreeformDrawingPrecision**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **FreeformDrawingPrecision** property is equivalent to setting the **Precision** option on the **Advanced** tab in the **Visio Options** dialog box (click the **File** tab, and then click **Options**).

Possible values for the  **FreeformDrawingPrecision** property range from 0 ( **Tight**) to 10 ( **Loose**). The default is 5.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **FreeformDrawingPrecision** property to print the current freeform drawing precision setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub FreeformDrawingPrecision_Example() 
 
    Dim vsoApplicationSettings As Visio.ApplicationSettings 
    Dim lngPrecisionSetting As Long 
 
    Set vsoApplicationSettings = Visio.Application.Settings 
    lngPrecisionSetting = vsoApplicationSettings.FreeformDrawingPrecision 
 
    Debug.Print lngPrecisionSetting 
 
End Sub
```


