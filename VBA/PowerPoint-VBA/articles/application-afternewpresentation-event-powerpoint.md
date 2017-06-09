---
title: Application.AfterNewPresentation Event (PowerPoint)
keywords: vbapp10.chm621020
f1_keywords:
- vbapp10.chm621020
ms.prod: powerpoint
api_name:
- PowerPoint.Application.AfterNewPresentation
ms.assetid: d95bb247-2ebd-263f-d6b5-9918204b9130
ms.date: 06/08/2017
---


# Application.AfterNewPresentation Event (PowerPoint)

Occurs after a presentation is created.


## Syntax

 _expression_. **AfterNewPresentation**( **_Pres_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|Name of the presentation.|

## Example

This example uses the  **RGB** function to set the slide master background color for the new presentation to salmon pink, and then applies the third color scheme to the new presentation.


```vb
Private Sub App_AfterNewPresentation(ByVal Pres As Presentation)

    With Pres

        Set CS3 = .ColorSchemes(3)

        CS3.Colors(ppBackground).RGB = RGB(240, 115, 100)

        .SlideMaster.ColorScheme = CS3

    End With

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

