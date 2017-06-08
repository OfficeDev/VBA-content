---
title: ShapeRange.TextEffect Property (Project)
ms.prod: project-server
ms.assetid: 83c2ca99-7ae1-0a93-41f1-2e53379b54ec
ms.date: 06/08/2017
---


# ShapeRange.TextEffect Property (Project)
Gets text formatting properties for the shape range. Read-only  **[TextEffectFormat](http://msdn.microsoft.com/en-us/library/office/ff834714%28v=office.15%29)**.

## Syntax

 _expression_. **TextEffect**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

The following example creates a shape range that contains a text box shape, sets the foreground color of text in the text frame to red, sets the foreground color of the text box shape to a yellowish tan, and then uses the  **TextEffect** property to set font properties.

If there were more than one text box shape in the shape range, the font properties of each text box would be changed accordingly.




```vb
Sub FormatTextBox()
    Dim theReport As Report
    Dim textShape As shape
    Dim reportName As String
    Dim sRange As ShapeRange
    
    reportName = "Textbox range report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set textShape = theReport.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 50, 350, 80)
    textShape.Name = "My text box"
    
    textShape.TextFrame2.TextRange.Text = "This is a test. It is only a test. "
    textShape.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = &;H2020CC
    textShape.Fill.ForeColor.RGB = &;H88CCCC
    
    Set sRange = theReport.Shapes.Range(Array("My text box"))
    
    With sRange.TextEffect
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


[ShapeRange Object](shaperange-object-project.md)
[Shape.TextEffect Property](shape-texteffect-property-project.md)
[TextEffectFormat](http://msdn.microsoft.com/en-us/library/office/ff834714%28v=office.15%29)
