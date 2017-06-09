---
title: Series.PictureType Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Series.PictureType
ms.assetid: 106933a2-49a7-e9d3-e5fa-fd2d0ab8974a
ms.date: 06/08/2017
---


# Series.PictureType Property (PowerPoint)

Returns or sets a value that specifies how pictures are displayed on a column or bar picture chart. Read/write  **[XlChartPictureType](xlchartpicturetype-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **PictureType**

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


## Example

The following example sets series one of the first chart in the active document to stretch pictures. You should run the example on a 2-D column chart that has picture data markers.




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).PictureType = xlStretch

    End If

End With
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

