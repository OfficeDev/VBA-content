---
title: PublishObject.Application Property (PowerPoint)
keywords: vbapp10.chm635001
f1_keywords:
- vbapp10.chm635001
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.Application
ms.assetid: fa4ed2eb-f1d0-da75-d0d0-3b9255d66cb6
ms.date: 06/08/2017
---


# PublishObject.Application Property (PowerPoint)

Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **PublishObject** object.


### Return Value

Application


## Example

In this example, a  **[Presentation](presentation-object-powerpoint.md)** object is passed to the procedure. The procedure adds a slide to the presentation and then saves the presentation in the folder where Microsoft PowerPoint is running.


```vb
Sub AddAndSave(pptPres As Presentation)

    pptPres.Slides.Add 1, 1

    pptPres.SaveAs pptPres.Application.Path &; "\Added Slide"

End Sub
```

This example displays the name of the application that created each linked OLE object on slide one in the active presentation.




```vb
For Each shpOle In ActivePresentation.Slides(1).Shapes

    If shpOle.Type = msoLinkedOLEObject Then

        MsgBox shpOle.OLEFormat.Application.Name

    End If

Next
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-powerpoint.md)

