---
title: ApplicationSettings.SnapStrengthExtensionsX Property (Visio)
keywords: vis_sdr.chm16251595
f1_keywords:
- vis_sdr.chm16251595
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthExtensionsX
ms.assetid: 45fb7005-34af-860f-ea59-a48e5a0b7a01
ms.date: 06/08/2017
---


# ApplicationSettings.SnapStrengthExtensionsX Property (Visio)

Specifies the distance in pixels along the  _x_ -axis that shape extension lines pull when snapping is enabled. Read/Write.


## Syntax

 _expression_ . **SnapStrengthExtensionsX**

 _expression_ A variable that represents a **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **SnapStrengthExtensionsX** property is equivalent to setting the **Extensions** option under **Snap strength** on the **Advanced** tab in the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthExtensionsX** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 13.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthExtensionsX** property to print the current snap strength extensions _x_ -axis setting in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthExtensionsX_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Dim lngSnapStrength As Long 
 
 Set vsoApplicationSettings = Visio.Application.Settings 
 lngSnapStrength = vsoApplicationSettings.SnapStrengthExtensionsX 
 
 Debug.Print lngSnapStrength 
 
End Sub
```


