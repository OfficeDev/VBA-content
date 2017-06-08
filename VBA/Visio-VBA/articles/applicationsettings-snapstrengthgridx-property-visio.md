---
title: ApplicationSettings.SnapStrengthGridX Property (Visio)
keywords: vis_sdr.chm16251575
f1_keywords:
- vis_sdr.chm16251575
ms.prod: visio
api_name:
- Visio.ApplicationSettings.SnapStrengthGridX
ms.assetid: ebe2489d-6643-4303-30fd-720446a4e19d
ms.date: 06/08/2017
---


# ApplicationSettings.SnapStrengthGridX Property (Visio)

Specifies the distance in pixels along the x-axis that gridlines pull when snapping is enabled. Read/write.


## Syntax

 _expression_ . **SnapStrengthGridX**

 _expression_ A variable that represents a **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **SnapStrengthGridX** property is equivalent to setting the **Grid** option under **Snap strength** on the **Advanced** tab in the **Snap &; Glue** dialog box (click the **Visual Aids** arrow on the **View** tab). Setting snap strength in the UI sets both _x_ and _y_ values to the same value.

The minimum allowable value for the  **SnapStrengthGridX** property is 0 (zero), and the maximum is 999. Attempting to set a value outside that range returns an error. The default value is 5.


## Example

ThisMicrosoft Visual Basic for Applications (VBA) macro shows how to use the  **SnapStrengthGridX** property to print the current snap strength grid _x_ -axis setting in the **Immediate** window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub SnapStrengthGridX_Example() 
 
    Dim vsoApplicationSettings As Visio.ApplicationSettings 
    Dim lngSnapStrength As Long 
 
    Set vsoApplicationSettings = Visio.Application.Settings 
    lngSnapStrength = vsoApplicationSettings.SnapStrengthGridX 
 
    Debug.Print lngSnapStrength 
 
End Sub
```


