---
title: ApplicationSettings.SnapStrengthRulerY Property (Visio)
keywords: vis_sdr.chm16251540
f1_keywords:
- vis_sdr.chm16251540
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthRulerY
ms.assetid: b0b6a3da-a87d-496e-901c-e6850e6c612b
ms.date: 06/08/2017
---


# ApplicationSettings.SnapStrengthRulerY Property (Visio)

Specifies the distance in pixels along the y-axis that rulers pull when snapping is enabled. Read/write.


## Syntax

 _expression_ . **SnapStrengthRulerY**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **SnapStrengthRulerY** property is equivalent to setting the **Rulers** option under **Snap strength** on the **Advanced** tab in the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthRulerY** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 4.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthRulerY** property to print the current snap strength ruler _y_ -axis setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthRulerY_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthRulerY 
 
 Debug.Print lngSnapStrength 
 
End Sub
```


