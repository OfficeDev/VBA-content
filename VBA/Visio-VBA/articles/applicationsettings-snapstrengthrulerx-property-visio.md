---
title: ApplicationSettings.SnapStrengthRulerX Property (Visio)
keywords: vis_sdr.chm16251545
f1_keywords:
- vis_sdr.chm16251545
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthRulerX
ms.assetid: 594b4730-94ac-de20-12df-97ae0df4b7f6
ms.date: 06/08/2017
---


# ApplicationSettings.SnapStrengthRulerX Property (Visio)

Specifies the distance in pixels along the x-axis that rulers pull when snapping is enabled. Read/write.


## Syntax

 _expression_ . **SnapStrengthRulerX**

 _expression_ A variable that represents a **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **SnapStrengthRulerX** property is equivalent to setting the **Rulers** option under **Snap strength** on the **Advanced** tab in the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthRulerX** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 4.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthRulerX** property to print the current snap strength ruler _x_ -axis setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthRulerX_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthRulerX 
 
 Debug.Print lngSnapStrength 
 
End Sub
```


