---
title: Shape.TextEffect Property (Project)
ms.prod: project-server
ms.assetid: 12fa0951-e3a5-807e-bebb-bff82650d200
ms.date: 06/08/2017
---


# Shape.TextEffect Property (Project)
Gets text formatting properties for the shape. Read-only  **[TextEffectFormat](http://msdn.microsoft.com/en-us/library/office/ff834714%28v=office.15%29)**.

## Syntax

 _expression_. **TextEffect**

 _expression_ A variable that represents a **Shape** object.


## Example

The following example sets the foreground color of text in a text frame to red, the foreground color of the text box shape to a yellowish tan, and then uses the  **TextEffect** property to set font properties.


```vb
Sub FormatTextBox()
    Dim theReport As Report
    Dim textShape As shape
    Dim reportName As String
    
    reportName = "Textbox report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set textShape = theReport.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 50, 350, 80)
    
    textShape.TextFrame2.TextRange.Text = "This is a test. It is only a test. "
    textShape.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = &;H2020CC
    textShape.Fill.ForeColor.RGB = &;H88CCCC
    
    With textShape.TextEffect
        .FontName = "Courier New"
        .FontBold = True
        .FontItalic = True
        .FontSize = 28
    End With
End Sub
```


## Property value

 **TEXTEFFECTFORMAT**


## See also


#### Other resources


[Shape Object](shape-object-project.md)
[ShapeRange.TextEffect Property](shaperange-texteffect-property-project.md)
[TextEffectFormat](http://msdn.microsoft.com/en-us/library/office/ff834714%28v=office.15%29)
