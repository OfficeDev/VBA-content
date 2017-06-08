---
title: Shapes.AddLabel Method (Project)
ms.prod: project-server
ms.assetid: 3fd21dbc-51b7-0e22-8c8a-359b1717932f
ms.date: 06/08/2017
---


# Shapes.AddLabel Method (Project)
Creates a label in a report, and returns a  **Shape** object that represents the new label.

## Syntax

 _expression_. **AddLabel** _(Orientation,_ _Left,_ _Top,_ _Width,_ _Height)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required|**MsoTextOrientation**|The text orientation within the label.|
| _Left_|Required|**Single**|The position (in points) of the left edge of the label relative to the left side of the report.|
| _Top_|Required|**Single**|The position (in points) of the top edge of the label relative to the top of the report.|
| _Width_|Required|**Single**|The width of the label, in points.|
| _Height_|Required|**Single**|The height of the label, in points.|
| _Orientation_|Required|MSOTEXTORIENTATION||
| _Left_|Required|FLOAT||
| _Top_|Required|FLOAT||
| _Width_|Required|FLOAT||
| _Height_|Required|FLOAT||

### Return value

 **Shape**


## Example

The following example adds a green label with the text "Hello report!" to a new report.


```vb
Sub AddHelloLabel()
    Dim shapeReport As Report
    Dim reportName As String
    Dim labelShape As shape
    
    ' Add a report.
    reportName = "Label report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)

    Set labelShape = shapeReport.Shapes.AddLabel(msoTextOrientationHorizontal, 30, 30, 120, 40)

    With labelShape
        With .Fill
            .BackColor.RGB = RGB(red:=&;H20, green:=&;HFF, blue:=&;H20)
            .Visible = msoTrue
        End With
        
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        
        With .TextFrame2.TextRange
            .Text = "Hello report!"
            .Font.Bold = msoTrue
            .Font.Name = "Calibri"
            .Font.Size = 18
        End With
    End With
End Sub
```


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
