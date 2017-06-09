---
title: OLEFormat.ObjectVerbs Property (PowerPoint)
keywords: vbapp10.chm562003
f1_keywords:
- vbapp10.chm562003
ms.prod: powerpoint
api_name:
- PowerPoint.OLEFormat.ObjectVerbs
ms.assetid: 895becb3-de86-638c-88e9-b9e72b6c713e
ms.date: 06/08/2017
---


# OLEFormat.ObjectVerbs Property (PowerPoint)

Returns a  **[ObjectVerbs](objectverbs-object-powerpoint.md)** collection that contains all the OLE verbs for the specified OLE object. Read-only.


## Syntax

 _expression_. **ObjectVerbs**

 _expression_ A variable that represents an **OLEFormat** object.


### Return Value

ObjectVerbs


## Example

This example displays all the available verbs for the OLE object contained in shape one on slide two in the active presentation. For this example to work, shape one must be a shape that represents an OLE object.


```vb
With ActivePresentation.Slides(2).Shapes(1).OLEFormat

    For Each v In .ObjectVerbs

        MsgBox v

    Next

End With
```

This example specifies that the OLE object represented by shape one on slide two in the active presentation will open when it is clicked during a slide show if "Open" is one of the OLE verbs for that object. For this example to work, shape one must be a shape that represents an OLE object.




```vb
With ActivePresentation.Slides(2).Shapes(1)

    For Each sVerb In .OLEFormat.ObjectVerbs

        If sVerb = "Open" Then

            With .ActionSettings(ppMouseClick)

                .Action = ppActionOLEVerb

                .ActionVerb = sVerb

            End With

            Exit For

        End If

    Next

End With
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-powerpoint.md)

