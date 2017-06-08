---
title: ApplicationSettings.SnapStrengthGuidesY Property (Visio)
keywords: vis_sdr.chm16251560
f1_keywords:
- vis_sdr.chm16251560
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthGuidesY
ms.assetid: 64d2c688-d900-c5e7-28c7-a0c24dcc854a
ms.date: 06/08/2017
---


# ApplicationSettings.SnapStrengthGuidesY Property (Visio)

Specifies the distance in pixels along the y-axis that guides pull when snapping is enabled. Read/write.


## Syntax

 _expression_ . **SnapStrengthGuidesY**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **SnapStrengthGuidesY** property is equivalent to setting the **Guides** option under **Snap strength** on the **Advanced** tab in the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthGuidesY** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 10.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthGuidesY** property to print the current snap strength guides _y_ -axis setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthGuidesY_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthGuidesY 
 
 Debug.Print lngSnapStrength 
 
End Sub
```


