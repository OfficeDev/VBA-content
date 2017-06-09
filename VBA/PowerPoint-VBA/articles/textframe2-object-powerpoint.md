---
title: TextFrame2 Object (PowerPoint)
keywords: vbapp10.chm678000
f1_keywords:
- vbapp10.chm678000
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2
ms.assetid: ae017598-8330-4673-db1a-53b284acb709
ms.date: 06/08/2017
---


# TextFrame2 Object (PowerPoint)

Represents the text frame in a  **[Shape](shape-object-powerpoint.md)** or **[ShapeRange](shaperange-object-powerpoint.md)** object. Contains the text in the text frame and exposes properties and methods that control the alignment and anchoring of the text frame.


## Remarks

Use the  **TextFrame2** property of the **Shape** and **ShapeRange** objects to return a **TextFrame2** object.

Use the  **HasTextFrame** property to determine whether a shape or shape range has a text frame, and use the **HasText** property to determine whether the text frame contains text.


## Example

The following example adds a rectangle to a slide, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Public Sub TextFrame2_Example()



    Set pptSlide = ActivePresentation.Slides(1)

    With pptSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2

        .TextRange.Text = "Here is some sample text"

        .MarginBottom = 10

        .MarginLeft = 10

        .MarginRight = 10

        .MarginTop = 10

    End With

    

End Sub
```

The following example shows how to use the  **HasTextFrame** property to determine whether a shape has a text frame, and then how to use the **HasText** property to determine whether the text frame contains text.




```vb
Public Sub HasTextFrame_Example()



    Set pptSlide = ActivePresentation.Slides(1)

    For Each pptShape In pptSlide.Shapes

        If pptShape.HasTextFrame Then

            With pptShape.TextFrame2

                If .HasText Then MsgBox .TextRange.Text

            End With

        End If

    Next

    

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

