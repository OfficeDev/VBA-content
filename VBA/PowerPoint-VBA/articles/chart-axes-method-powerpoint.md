---
title: Chart.Axes Method (PowerPoint)
keywords: vbapp10.chm684016
f1_keywords:
- vbapp10.chm684016
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Axes
ms.assetid: 6f740a9e-2baa-5a84-ea51-6a39452e227e
ms.date: 06/08/2017
---


# Chart.Axes Method (PowerPoint)

Returns a collection of axes on the chart.


## Syntax

 _expression_. **Axes**( **_Type_**, **_AxisGroup_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters



| <strong>Name</strong> | <strong>Required/Optional</strong> | <strong>Data Type</strong>                                                                                                                              | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                           |
|:----------------------|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>Type</em>         | Optional                           | <strong>Variant</strong>                                                                                                                                | The axis to return. Can be one of the following  <strong><a href="xlaxistype-enumeration-powerpoint.md" data-raw-source="[XlAxisType](xlaxistype-enumeration-powerpoint.md)">XlAxisType</a></strong> constants: <strong>xlValue</strong>, <strong>xlCategory</strong>, or <strong>xlSeriesAxis</strong> ( <strong>xlSeriesAxis</strong> is valid only for 3-D charts). |
| <em>AxisGroup</em>    | Optional                           | <strong><a href="xlaxisgroup-enumeration-powerpoint.md" data-raw-source="[XlAxisGroup](xlaxisgroup-enumeration-powerpoint.md)">XlAxisGroup</a></strong> | One of the enumeration values that specifies the axis group. The default is  <strong>xlPrimary</strong>.                                                                                                                                                                                                                                                               |

 **Note**  3-D charts have only one axis group.

|

### Return Value

An [Axes](axes-object-powerpoint.md) object that contains the selected axes from the chart.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds an axis label to the category axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        With .Chart.Axes(xlCategory) 
            .HasTitle = True 
            .AxisTitle.Text = "July Sales" 
        End With 
    End If 
End With
```

The following example turns off major gridlines in the category axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        .Chart.Axes(xlCategory). _ 
            HasMajorGridlines = False 
    End If 
End With
```

The following example turns off all gridlines for all axes in the first chart of the active document.




```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        For Each a In .Chart.Axes 
            a.HasMajorGridlines = False 
            a.HasMinorGridlines = False 
        Next 
    End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

