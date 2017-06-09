---
title: ShapeRange.Value Property (Project)
ms.prod: project-server
ms.assetid: 19793067-571a-38b9-30b0-7b84b0864290
ms.date: 06/08/2017
---


# ShapeRange.Value Property (Project)
Gets an individual  **Shape** object in the **ShapeRange** collection. Read-only **Shape**.

## Syntax

 _expression_. **Value**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

 **Value** is the default property for a **ShapeRange** object.


## Example

The following example creates a report named "Test Report", creates two shapes, and then adds the shapes to a  **ShapeRange** object. The statement that begins with `sRange.Value(1)` gets the first shape in the shape range. The statement that begins with `sRange(2)` invokes the default **Value** property and gets the second shape in the shape range.


```vb
Sub TestShapeRangeValue()
    Dim theReport As Report
    Dim textShape1 As shape
    Dim textShape2 As shape
    Dim reportName As String
    Dim sRange As ShapeRange
    
    reportName = "Test Report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set textShape1 = theReport.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 50, 350, 80)
    textShape1.Name = "Text box 1"
    
    Set textShape2 = theReport.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 130, 350, 80)
    textShape2.Name = "Text box 2"
    
    Set sRange = theReport.Shapes.Range(Array("Text box 1", "Text box 2"))
    
    sRange.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        
    sRange.Value(1).TextFrame2.TextRange.Text = "This is a test. It is only a test."
    sRange(2).TextFrame2.TextRange.Text = "This is text box 2."
End Sub
```


## Property value

 **SHAPE**


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
[Shape Object](shape-object-project.md)
