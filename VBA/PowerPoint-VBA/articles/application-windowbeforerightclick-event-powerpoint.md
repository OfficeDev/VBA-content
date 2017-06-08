---
title: Application.WindowBeforeRightClick Event (PowerPoint)
keywords: vbapp10.chm621002
f1_keywords:
- vbapp10.chm621002
ms.prod: powerpoint
api_name:
- PowerPoint.Application.WindowBeforeRightClick
ms.assetid: e6239915-f487-3619-c84f-d436d645e6c0
ms.date: 06/08/2017
---


# Application.WindowBeforeRightClick Event (PowerPoint)

Occurs when you right-click a shape, a slide, a notes page, or some text. This event is triggered by the  **MouseUp** event.


## Syntax

 _expression_. **WindowBeforeRightClick**( **_Sel_**, **_Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sel_|Required|**Selection**|The selection below the mouse pointer when the right-click occurred.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the default context menu does not appear when the procedure is finished.|

## Example

This example creates a duplicate of the selected shape. If the shape has a text frame, it adds the text "Duplicate Shape" to the new shape. Setting the Cancel argument to  **True** then prevents the default context menu from appearing.


```vb
Private Sub App_WindowBeforeRightClick(ByVal Sel As Selection, ByVal Cancel As Boolean)

    With ActivePresentation.Selection.ShapeRange

        If .HasTextFrame Then

            .Duplicate.TextFrame.TextRange.Text = "Duplicate Shape"

        Else

            .Duplicate

        End If

        Cancel = True

    End With

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

