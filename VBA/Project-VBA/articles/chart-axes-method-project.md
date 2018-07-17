---
title: Chart.Axes Method (Project)
ms.prod: project-server
ms.assetid: 0ab295f0-de68-7b8f-50a7-55a1e378080b
ms.date: 06/08/2017
---


# Chart.Axes Method (Project)
Returns an object that represents either a single axis or a collection of the axes on the chart.

## Syntax

 _expression_. **Axes** _(Type,_ _AxisGroup)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Variant**|Specifies the axis to return. Can be one of the following  **Office.XlAxisType** constants: **xlValue**,  **xlCategory**, or  **xlSeriesAxis** ( **xlSeriesAxis** is valid only for 3-D charts).|
| _AxisGroup_|Optional|**Office.XlAxisGroup**|Specifies the axis group. The default value is  **xlPrimary**; that is, if the  _AxisGroup_ argument is omitted, the primary group is used. 3-D charts have only one axis group.|
| _Type_|Optional|VARIANT||
| _AxisGroup_|Optional|XLAXISGROUP||

### Return value

 **Object**


## Examples

The  **SetAxisTitle** macro adds an axis label to the category axis in the chart.


```vb
Sub SetAxisTitle()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    With chartShape.Chart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Task"
    End With
End Sub
```

The  **AddCategoryGridlines** macro adds gridlines to the category axis in the chart.




```vb
Sub AddCategoryGridlines()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.Axes(xlCategory).HasMajorGridlines = True
End Sub
```

The RemoveGridlines macro removes the major gridlines from both the category and value axes.




```vb
Sub RemoveGridlines()
    Dim chartShape As Shape
    Dim reportName As String
    Dim axes As Object
    Dim a As Object
    
    reportName = "Simple scalar chart"
    
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.Axes(xlCategory).HasMajorGridlines = False
    chartShape.Chart.Axes(xlValue).HasMajorGridlines = False
End Sub
```


