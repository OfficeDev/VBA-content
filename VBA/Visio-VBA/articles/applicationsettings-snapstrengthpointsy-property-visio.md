---
title: ApplicationSettings.SnapStrengthPointsY Property (Visio)
keywords: vis_sdr.chm16251550
f1_keywords:
- vis_sdr.chm16251550
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthPointsY
ms.assetid: 7719694e-993a-2792-3f6f-3d697ef34790
ms.date: 06/08/2017
---


# ApplicationSettings.SnapStrengthPointsY Property (Visio)

Specifies the distance in pixels along the y-axis that points pull when snapping is enabled. Read/write.


## Syntax

 _expression_ . **SnapStrengthPointsY**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **SnapStrengthPointsY** property is equivalent to setting the **Points** option under **Snap strength** on the **Advanced** tab in the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthPointsY** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 10.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthPointsY** property to print the current snap strength points _y_ -axis setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthPointsY_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthPointsY 
 
 Debug.Print lngSnapStrength 
 
End Sub
```


