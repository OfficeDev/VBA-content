---
title: Options Object (PowerPoint)
keywords: vbapp10.chm667000
f1_keywords:
- vbapp10.chm667000
ms.prod: powerpoint
api_name:
- PowerPoint.Options
ms.assetid: c129bafc-9927-0171-769e-21649ead7dca
ms.date: 06/08/2017
---


# Options Object (PowerPoint)

Represents application options in Microsoft PowerPoint.


## Example

Use the  **[Options](application-options-property-powerpoint.md)** property to return an **Options** object. The following example sets three application options for PowerPoint.


```vb
Sub TogglePasteOptionsButton()

    With Application.Options

        If .DisplayPasteOptions = False Then

            .DisplayPasteOptions = True

        End If

    End With

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

