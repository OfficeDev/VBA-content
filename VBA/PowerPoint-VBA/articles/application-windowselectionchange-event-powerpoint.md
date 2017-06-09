---
title: Application.WindowSelectionChange Event (PowerPoint)
keywords: vbapp10.chm621001
f1_keywords:
- vbapp10.chm621001
ms.prod: powerpoint
api_name:
- PowerPoint.Application.WindowSelectionChange
ms.assetid: 069f4afe-2302-28fa-4d86-57afe8c3c2ab
ms.date: 06/08/2017
---


# Application.WindowSelectionChange Event (PowerPoint)

Occurs when the selection of text, a shape, or a slide in the active document window changes, whether in the user interface or in code.


## Syntax

 _expression_. **WindowSelectionChange**( **_Sel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sel_|Required|**Selection**|Represents the object selected.|

## Example

This example determines when a different slide is being selected and changes the background color of the newly selected slide.


```vb
Private Sub App_WindowSelectionChange(ByVal Sel As Selection)

    With Sel
        If .Type = ppSelectionNone Then
            With .SlideRange(1)
                .ColorScheme.Colors(ppBackground).RGB = _
                    RGB(240, 115, 100)
            End With
        End If
    End With

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

