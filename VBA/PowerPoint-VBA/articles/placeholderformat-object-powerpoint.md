---
title: PlaceholderFormat Object (PowerPoint)
keywords: vbapp10.chm545000
f1_keywords:
- vbapp10.chm545000
ms.prod: powerpoint
api_name:
- PowerPoint.PlaceholderFormat
ms.assetid: 5e204d07-7ec0-b08c-497c-7f0174d28782
ms.date: 06/08/2017
---


# PlaceholderFormat Object (PowerPoint)

Contains properties that apply specifically to placeholders, such as placeholder type.


## Example

Use the [PlaceholderFormat](shape-placeholderformat-property-powerpoint.md)property to return a  **PlaceholderFormat** object. The following example adds text to placeholder one on slide one in the active presentation if that placeholder exists and is a horizontal title placeholder.


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
                    MsgBox "There's no horizontal " _
                        "title on this slide"
            End Select
        End With
    End If
End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

