---
title: Shape.PlaceholderFormat Property (PowerPoint)
keywords: vbapp10.chm547046
f1_keywords:
- vbapp10.chm547046
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.PlaceholderFormat
ms.assetid: 4ccd4f93-74fc-be23-5ef4-0089d7247724
ms.date: 06/08/2017
---


# Shape.PlaceholderFormat Property (PowerPoint)

Returns a  **[PlaceholderFormat](placeholderformat-object-powerpoint.md)** object that contains the properties that are unique to placeholders. Read-only.


## Syntax

 _expression_. **PlaceholderFormat**

 _expression_ A variable that represents a **Shape** object.


### Return Value

PlaceholderFormat


## Example

This example adds text to placeholder one on slide one in the active presentation if that placeholder is a horizontal title placeholder.


```vb
With ActivePresentation.Slides(1).Shapes.Placeholders
    If .Count > 0 Then
        With .Item(1)
            Select Case .PlaceholderFormat.Type
                Case ppPlaceholderTitle
                    .TextFrame.TextRange = "Title Text"

                Case ppPlaceholderCenterTitle
                    .TextFrame.TextRange = "Centered Title Text"

                Case Else
                    MsgBox "There's no horizontal" &; _
                        "title on this slide"
            End Select
        End With
    End If
End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

