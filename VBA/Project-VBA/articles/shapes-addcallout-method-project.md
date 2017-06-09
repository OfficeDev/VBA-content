---
title: Shapes.AddCallout Method (Project)
ms.prod: project-server
ms.assetid: 6c183677-d869-f493-7226-14cca4329aae
ms.date: 06/08/2017
---


# Shapes.AddCallout Method (Project)
Creates a borderless line callout in a report. Returns a  **Shape** object that represents the new callout.

## Syntax

 _expression_. **AddCallout** _(Type,_ _Left,_ _Top,_ _Width,_ _Height)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**MsoCalloutType**|The type of callout.|
| _Left_|Required|**Single**|The position, in points, of the left edge of the bounding box for the callout.|
| _Top_|Required|**Single**|The position, in points, of the top edge of the bounding box for the callout.|
| _Width_|Required|**Single**|The width, in points, of the bounding box for the callout.|
| _Height_|Required|**Single**|The height, in points, of the bounding box for the callout.|
| _Type_|Required|MSOCALLOUTTYPE||
| _Left_|Required|FLOAT||
| _Top_|Required|FLOAT||
| _Width_|Required|FLOAT||
| _Height_|Required|FLOAT||

### Return value

 **Shape**


### Remarks

The  _Type_ parameter can be one of the following **MsoCalloutType** constants:


||
|:-----|
|**msoCalloutOne**: A single-segment callout line that can be horizontal or vertical.|
|**msoCalloutTwo**: A single-segment callout line that rotates freely.|
|**msoCalloutMixed**: A return value that indicates a combination of the other states.|
|**msoCalloutThree**: A two-segment line, where the segment ends can be dragged to different positions.|
|**msoCalloutFour**: A three-segment line.|
You can insert a greater variety of callouts, such as balloons and clouds, by using the  **[AddShape](shapes-addshape-method-project.md)** method.


### Example

The following example adds a callout with a two-segment callout line, sets the angle of the end segment to 60 degrees from the vertical, and adds text to the callout.


```vb
Sub AddCallout()
    Dim oReports As Reports
    Dim oReport As Report
    Dim calloutShape As shape
    Dim reportName As String
    
    reportName = "Report 1"
    Set oReports = ActiveProject.Reports

    If oReports.IsPresent(reportName) Then
        ' Make the report the active view.
        oReports(reportName).Apply
        
        Set oReport = oReports(reportName)
        
        Set calloutShape = oReport.Shapes.AddCallout(Type:=msoCalloutTwo, _
                                        left:=200, top:=5, width:=100, height:=50)
        With calloutShape
            .Callout.Type = msoCalloutThree
            .Callout.Angle = msoCalloutAngle60
            .BackgroundStyle = msoBackgroundStylePreset10
            .TextFrame2.TextRange.Text = "This is a test"
        End With
    Else
         MsgBox Prompt:="The requested report, '" &; reportName _
            &; "', does not exist.", Title:="Report error"
    End If
End Sub
```


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
[AddShape Method](shapes-addshape-method-project.md)
