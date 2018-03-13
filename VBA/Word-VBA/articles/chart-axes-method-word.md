---
title: Chart.Axes Method (Word)
ms.prod: word
api_name:
- Word.Chart.Axes
ms.assetid: 37f422b5-31f2-92ce-c04e-a837b0a3d407
ms.date: 06/08/2017
---


# Chart.Axes Method (Word)

Returns a collection of axes on the chart.


## Syntax

 _expression_ . **Axes**( **_Type_** , **_AxisGroup_** )

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


### Parameters



| <strong>Name</strong> | <strong>Required/Optional</strong> | <strong>Data Type</strong>                                                                                                                  | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                 |
|:----------------------|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>Type</em>         | Optional                           | <strong>Variant</strong>                                                                                                                    | The axis to return. Can be one of the following  <strong><a href="xlaxistype-enumeration-word.md" data-raw-source="[XlAxisType](xlaxistype-enumeration-word.md)">XlAxisType</a></strong> constants: <strong>xlValue</strong> , <strong>xlCategory</strong> , or <strong>xlSeriesAxis</strong> ( <strong>xlSeriesAxis</strong> is valid only for 3-D charts). |
| <em>AxisGroup</em>    | Optional                           | <strong><a href="xlaxisgroup-enumeration-word.md" data-raw-source="[XlAxisGroup](xlaxisgroup-enumeration-word.md)">XlAxisGroup</a></strong> | One of the enumeration values that specifies the axis group. The default is  <strong>xlPrimary</strong> .                                                                                                                                                                                                                                                    |

 **Note**  3-D charts have only one axis group.

|

### Return Value

An [Axes](axes-object-word.md) object that contains the selected axes from the chart.


## Example

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


[Chart Object](chart-object-word.md)

