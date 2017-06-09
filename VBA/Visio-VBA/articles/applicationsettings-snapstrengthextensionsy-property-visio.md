---
title: ApplicationSettings.SnapStrengthExtensionsY Property (Visio)
keywords: vis_sdr.chm16251590
f1_keywords:
- vis_sdr.chm16251590
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthExtensionsY
ms.assetid: 01540007-8cbb-e551-6917-85295c99185a
ms.date: 06/08/2017
---


# ApplicationSettings.SnapStrengthExtensionsY Property (Visio)

Specifies the distance in pixels along the  _y-_ axis that shape extension lines pull when snapping is enabled. Read/write.


## Syntax

 _expression_ . **SnapStrengthExtensionsY**

 _expression_ A variable that represents a **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **SnapStrengthExtensionsY** property is equivalent to setting the **Extensions** option under **Snap strength** on the **Advanced** tab in the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthExtensionsY** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 13.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthExtensionsY** property to print the current snap strength extensions _y_ -axis setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthExtensionsY_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthExtensionsY 
 
 Debug.Print lngSnapStrength 
 
End Sub
```


