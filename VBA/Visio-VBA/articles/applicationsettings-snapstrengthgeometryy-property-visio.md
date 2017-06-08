---
title: ApplicationSettings.SnapStrengthGeometryY Property (Visio)
keywords: vis_sdr.chm16251580
f1_keywords:
- vis_sdr.chm16251580
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthGeometryY
ms.assetid: 8e5b3bf3-4cb6-af1c-1812-863c247608b9
ms.date: 06/08/2017
---


# ApplicationSettings.SnapStrengthGeometryY Property (Visio)

Specifies the distance in pixels along the  _y_ -axis that shape geometry pulls when snapping is enabled. Read/write.


## Syntax

 _expression_ . **SnapStrengthGeometryY**

 _expression_ A variable that represents a **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **SnapStrengthGeometryY** property is equivalent to setting the **Geometry** option under **Snap strength** on the **Advanced** tab in the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthGeometryY** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 8.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthGeometryY** property to print the current snap strength geometry _y_ -axis setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthGeometryY_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthGeometryY 
 
 Debug.Print lngSnapStrength 
 
End Sub
```


