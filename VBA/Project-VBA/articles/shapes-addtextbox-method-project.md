---
title: Shapes.AddTextbox Method (Project)
ms.prod: project-server
ms.assetid: ee8c619f-8b35-6f94-e680-86dbeedd6d19
ms.date: 06/08/2017
---


# Shapes.AddTextbox Method (Project)
Adds a text box to the report, and returns a  **Shape** object that represents the new text box.

## Syntax

 _expression_. **AddTextbox** _(Orientation,_ _Left,_ _Top,_ _Width,_ _Height)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required|**MsoTextOrientation**|The orientation of the text box. Some constants may not be available, depending on the language that is installed.|
| _Left_|Required|**Single**|The position, in points, of the left edge of the text box.|
| _Top_|Required|**Single**|The position, in points, of the top edge of the text box.|
| _Width_|Required|**Single**|The width, in points, of the text box.|
| _Height_|Required|**Single**|The height, in points, of the text box.|
| _Orientation_|Required|MSOTEXTORIENTATION||
| _Left_|Required|FLOAT||
| _Top_|Required|FLOAT||
| _Width_|Required|FLOAT||
| _Height_|Required|FLOAT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Shape**


## Example

The following example adds a text box with a light yellow background and a visible border. The text string is formatted and manipulated by using members of the  **TextFrame2** object.


```vb
Sub AddTextBoxShape()
    Dim theReport As Report
    Dim textShape As shape
    Dim reportName As String
    
    reportName = "Textbox report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set textShape = theReport.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 50, 300, 100)
    
    textShape.TextFrame2.TextRange.Characters.Text = "This is a test. It is only a test. " _
        &; "If it had been real information, there would be some real text here."
    textShape.TextFrame2.TextRange.Characters(1, 15).ParagraphFormat.FirstLineIndent = 10
    textShape.TextFrame2.TextRange.Characters(16).InsertBefore vbCrLf
    
    ' Set the font for the first 15 characters to dark blue bold.
    With textShape.TextFrame2.TextRange.Characters(1, 15).Font
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent5
        .Fill.Solid
        .Fill.Visible = msoTrue
        .Size = 14
        .Bold = msoTrue
    End With

    With textShape.Fill
        .ForeColor.RGB = RGB(255, 255, 160)
        .Visible = msoTrue
    End With
   
    With textShape.Line
        .Weight = 1
        .Visible = msoTrue
    End With
End Sub
```


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
[TextFrame2 Property](shape-textframe2-property-project.md)
[MsoTextOrientation Enumeration (Office)](http://msdn.microsoft.com/en-us/library/office/ff862778%28v=office.15%29)
