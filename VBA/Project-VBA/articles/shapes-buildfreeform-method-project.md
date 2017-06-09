---
title: Shapes.BuildFreeform Method (Project)
ms.prod: project-server
ms.assetid: 257f76e3-3b37-5b58-cb78-f6fcebe1ca29
ms.date: 06/08/2017
---


# Shapes.BuildFreeform Method (Project)
Creates a  **FreeformBuilder** object that represents a new freeform drawing. The freeform drawing can be converted into a **Shape** object.

## Syntax

 _expression_. **BuildFreeform** _(EditingType,_ _X1,_ _Y1)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EditingType_|Required|**MsoEditingType**|The editing property of the first node.|
| _X1_|Required|**Single**|The position (in points) of the first node in the freeform drawing, relative to the left edge of the report.|
| _Y1_|Required|**Single**|The position (in points) of the first node in the freeform drawing, relative to the top edge of the report.|
| _EditingType_|Required|MSOEDITINGTYPE||
| _X1_|Required|FLOAT||
| _Y1_|Required|FLOAT||
|Name|Required/Optional|Data type|Description|

### Return value

 **FreeformBuilder**


## Remarks

Use the  **AddNodes** method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the **ConvertToShape** method to convert the **FreeformBuilder** object into a **Shape** object that has the geometric description that you defined.


## Example

The following example adds a freeform with five vertices to the report, converts the freeform to a shape, and then changes the background style of the shape.


```vb
Sub AddFreeform2()
    Dim shapeReport As Report
    Dim reportName As String
    Dim freeformBuild As FreeformBuilder
    Dim freeformShape As shape

    reportName = "Freeform2 report"
    Set shapeReport = ActiveProject.Reports.Add(reportName)
    
    Set freeformBuild = shapeReport.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    
    With freeformBuild
        .AddNodes msoSegmentCurve, msoEditingCorner, 380, 230, 400, 450, 300
        .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200
        .AddNodes msoSegmentLine, msoEditingAuto, 480, 400
        .AddNodes msoSegmentLine, msoEditingAuto, 360, 200
        .ConvertToShape
    End With
    
    Set freeformShape = shapeReport.Shapes(1)
    
    freeformShape.BackgroundStyle = msoBackgroundStylePreset10
End Sub
```


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
