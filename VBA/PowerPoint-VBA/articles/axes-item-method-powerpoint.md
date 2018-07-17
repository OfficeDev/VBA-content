---
title: Axes.Item Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Axes.Item
ms.assetid: 61657765-2c92-5fdf-c3a9-0c75ca70fe68
ms.date: 06/08/2017
---


# Axes.Item Method (PowerPoint)

Returns a single  **[Axis](axis-object-powerpoint.md)** object from an **Axes** collection.


## Syntax

 _expression_. **Item**( **_Type_**, **_AxisGroup_** )

 _expression_ A variable that represents an **[Axes](axes-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[XlAxisType](xlaxistype-enumeration-powerpoint.md)**|The axis type.|
| _AxisGroup_|Optional|**[XlAxisGroup](xlaxisgroup-enumeration-powerpoint.md)**|The axis.|

## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the title text for the category axis for the first chart in the active document.




```vb
With ActivePresentation.Slides(1).Shapes(1)

    If .HasChart Then

        With .Chart.Axes.Item(xlCategory)

            .HasTitle = True

            .AxisTitle.Caption = "1994"

        End With

    End If

End With
```


## See also


#### Concepts


[Axes Object](axes-object-powerpoint.md)

