---
title: TextFrame Object (PowerPoint)
keywords: vbapp10.chm558000
f1_keywords:
- vbapp10.chm558000
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame
ms.assetid: 03346e81-71b2-0b9e-843d-fb8aa0e3c868
ms.date: 06/08/2017
---


# TextFrame Object (PowerPoint)

Represents the text frame in a  **Shape** object. Contains the text in the text frame and the properties and methods that control the alignment and anchoring of the text frame.


## Example

Use the  **TextFrame** property to return a **TextFrame** object. The following example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _

        .AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

    .TextRange.Text = "Here is some test text"

    .MarginBottom = 10

    .MarginLeft = 10

    .MarginRight = 10

    .MarginTop = 10

End With
```

Use the [HasTextFrame](http://msdn.microsoft.com/library/ea1a53e4-32d8-e51f-9e60-9ef719c0d973%28Office.15%29.aspx)property to determine whether a shape has a text frame, and use the [HasText](http://msdn.microsoft.com/library/7bce3bae-38e7-d9d4-b67c-9454fafc620f%28Office.15%29.aspx)property to determine whether the text frame contains text, as shown in the following example.




```
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.HasTextFrame Then

        With s.TextFrame

            If .HasText Then MsgBox .TextRange.Text

        End With

    End If

Next
```


## Methods



|**Name**|
|:-----|
|[DeleteText](http://msdn.microsoft.com/library/0971765b-8d2c-a34a-7184-119af42be835%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/18ee8f34-836e-0e56-7187-aa32be937036%28Office.15%29.aspx)|
|[AutoSize](http://msdn.microsoft.com/library/771f5217-abce-f70d-743d-e17532dabd9e%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/7e198a9e-38eb-6f1a-38f6-e24bcac43190%28Office.15%29.aspx)|
|[HasText](http://msdn.microsoft.com/library/7bce3bae-38e7-d9d4-b67c-9454fafc620f%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/9f694882-ce8d-d412-d60e-5217e92a81a7%28Office.15%29.aspx)|
|[MarginBottom](http://msdn.microsoft.com/library/c1798b95-cb98-9dfd-5958-39fdea177b6e%28Office.15%29.aspx)|
|[MarginLeft](http://msdn.microsoft.com/library/c00a6b6c-0a67-5738-f31f-3714e2bf430d%28Office.15%29.aspx)|
|[MarginRight](http://msdn.microsoft.com/library/57ab53e7-1fbf-09b6-13c4-7cb0a814d9e3%28Office.15%29.aspx)|
|[MarginTop](http://msdn.microsoft.com/library/78ae54cd-1841-950b-c06e-c693fa5daebb%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/ce6a9578-3cbd-9b73-e374-c43fa4748054%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/3c5706c9-188e-6946-6e87-1501f32b1ce3%28Office.15%29.aspx)|
|[Ruler](http://msdn.microsoft.com/library/496ef8d2-b8c5-71a6-93d4-23e0a8d171f3%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/4a565e39-8bfa-7370-3ed6-57c442796144%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/fc38f7d8-25f7-5605-0f63-aa116fcf46d9%28Office.15%29.aspx)|
|[WordWrap](http://msdn.microsoft.com/library/f6077142-9afd-b274-7301-3e63d962e7b3%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
