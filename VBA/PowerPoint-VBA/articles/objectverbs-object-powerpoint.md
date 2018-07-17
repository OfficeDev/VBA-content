---
title: ObjectVerbs Object (PowerPoint)
keywords: vbapp10.chm564000
f1_keywords:
- vbapp10.chm564000
ms.prod: powerpoint
api_name:
- PowerPoint.ObjectVerbs
ms.assetid: 71dfd143-cec6-8b6f-7d0f-5229bc442d92
ms.date: 06/08/2017
---


# ObjectVerbs Object (PowerPoint)

Represents the collection of OLE verbs for the specified OLE object. OLE verbs are the operations supported by an OLE object. Commonly used OLE verbs are "play" and "edit."


## Example

Use the  **ObjectVerbs** property to return an **ObjectVerbs** object. The following example displays all the available verbs for the OLE object contained in shape one on slide two in the active presentation. For this example to work, shape one must contain an OLE object.


```vb
With ActivePresentation.Slides(2).Shapes(1).OLEFormat

    For Each v In .ObjectVerbs

        MsgBox v

    Next

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

