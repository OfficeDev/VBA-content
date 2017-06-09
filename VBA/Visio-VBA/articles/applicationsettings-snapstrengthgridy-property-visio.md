---
title: ApplicationSettings.SnapStrengthGridY Property (Visio)
keywords: vis_sdr.chm16251570
f1_keywords:
- vis_sdr.chm16251570
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthGridY
ms.assetid: 0fc60e09-0315-d981-7375-9c5fd71ec6bd
ms.date: 06/08/2017
---


# ApplicationSettings.SnapStrengthGridY Property (Visio)

Specifies the distance in pixels along the  _y_-axis that gridlines pull when snapping is enabled. Read/write.


## Syntax

 _expression_ . **SnapStrengthGridY**

 _expression_ A variable that represents a **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **SnapStrengthGridY** property is equivalent to setting the **Grid** option under **Snap strength** on the **Advanced** tab in the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthGridY** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 5.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthGridY** property to print the current snap strength grid _y_ -axis setting in the **Immediate** window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthGridY_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthGridY 
 
 Debug.Print lngSnapStrength 
 
End Sub
```


