---
title: ShapeRange.Select Method (Project)
ms.prod: project-server
ms.assetid: 41e923f7-a34f-d79a-e05c-55c8d0129ed5
ms.date: 06/08/2017
---


# ShapeRange.Select Method (Project)
Selects each shape in a shape range.

## Syntax

 _expression_. **Select** _(Replace)_

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Replace_|Optional|**Variant**|**True** replaces the current selection with the new selection. **False** adds the new selection to the current selection. The default value is **True**.|
| _Replace_|Optional|VARIANT||

### Return value

 **Nothing**


## Example

The following example creates three shapes, assigns two shapes to the first range, assigns the other shape to the second range, and then selects the shape ranges. Because the second range selection adds to the first range selection, all three shapes are selected (see Figure 1).


```vb
Sub SelectShapes()
    Dim theReport As Report
    Dim shp1 As shape
    Dim shp2 As shape
    Dim shp3 As shape
    Dim reportName As String
    Dim sRange1 As ShapeRange
    Dim sRange2 As ShapeRange
    
    reportName = "Select Report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set shp1 = theReport.Shapes.AddShape(msoShapeActionButtonHelp, 20, 50, 20, 30)
    Set shp2 = theReport.Shapes.AddShape(msoShapeBalloon, 100, 50, 30, 50)
    Set shp3 = theReport.Shapes.AddShape(msoShapeWave, 140, 50, 30, 50)
        
    Set sRange1 = theReport.Shapes.Range(Array(2, 3))
    Set sRange2 = theReport.Shapes.Range(1)
    
    sRange1.Select
    sRange2.Select False
End Sub
```


**Figure 1. Using the Select method to add to a selection**

![Using the Select method to add a selection](images/pj15_VBA_ShapeRange_Select.gif)


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
