---
title: ShapeRange.TextFrame2 Property (Project)
ms.prod: project-server
ms.assetid: 12cc5f21-09c5-adea-1253-40a6eaf17761
ms.date: 06/08/2017
---


# ShapeRange.TextFrame2 Property (Project)
Gets a  **TextFrame2** object that contains the text in a text frame and the members that control the alignment, anchoring, and other features of the text frame. Read-only **[TextFrame2](http://msdn.microsoft.com/en-us/library/office/ff822136%28v=office.15%29)**.

## Syntax

 _expression_. **TextFrame2**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

A  **TextFrame2** object contains many of the same properties as a **TextFrame** object, plus additional properties such as **AutoSize**,  **ThreeD**, and  **WordArtformat**.


## Example

The following example creates two text boxes and adds them to a  **ShapeRange** object, sets both text frames to automatically fit the text, sets the foreground color of text in the first text box shape to red, sets the foreground color of the shape range to a yellowish tan, and then uses the **TextEffect** property to set font properties on both text boxes in the shape range.

The  **TextFrame2** property for the **ShapeRange** object is shown in bold font.




```vb
Sub FormatTextBox()
    Dim theReport As Report
    Dim textShape1 As shape
    Dim textShape2 As shape
    Dim reportName As String
    Dim sRange As ShapeRange
    
    reportName = "Textbox range report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set textShape1 = theReport.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 50, 350, 80)
    textShape1.Name = "Text box 1"
    
    Set textShape2 = theReport.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 130, 350, 80)
    textShape2.Name = "Text box 2"
    
    Set sRange = theReport.Shapes.Range(Array("Text box 1", "Text box 2"))
        
    sRange.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        
    sRange(1).TextFrame2.TextRange.Text = "This is a test. It is only a test."
    sRange(2).TextFrame2.TextRange.Text = "This is text box 2."
    sRange(1).TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = &;H2020CC
    sRange.Fill.ForeColor.RGB = &;H88CCCC
        
    With sRange.TextEffect
        .FontName = "Courier New"
        .FontBold = True
        .FontItalic = True
        .FontSize = 28
    End With
    
    sRange(2).Select
End Sub
```


## Property value

 **TEXTFRAME2**


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
[TextFrame2](http://msdn.microsoft.com/en-us/library/office/ff822136%28v=office.15%29)
