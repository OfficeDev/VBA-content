---
title: DocumentWindows.Application Property (PowerPoint)
keywords: vbapp10.chm509001
f1_keywords:
- vbapp10.chm509001
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindows.Application
ms.assetid: 8a2cf838-7f79-d346-3a14-597c0f9fd4ef
ms.date: 06/08/2017
---


# DocumentWindows.Application Property (PowerPoint)

Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.


## Syntax

 _expression_. **Application**

 _expression_ A variable that represents a **DocumentWindows** object.


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


[DocumentWindows Object](documentwindows-object-powerpoint.md)


