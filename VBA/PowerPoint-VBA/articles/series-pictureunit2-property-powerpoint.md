---
title: Series.PictureUnit2 Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Series.PictureUnit2
ms.assetid: 83ccb10a-1883-9665-8a63-4494e853aa72
ms.date: 06/08/2017
---


# Series.PictureUnit2 Property (PowerPoint)

Returns or sets the unit for each picture on the chart if the  **[PictureType](series-picturetype-property-powerpoint.md)** property is set to **xlStackScale**; otherwise, this property is ignored. Read/write **Double**.


## Syntax

 _expression_. **PictureUnit2**

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets series one for the first chart in the active document to stack pictures and uses each picture to represent five units. You should run the example on a 2-D column chart that has picture data markers.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1)

            .PictureType = xlScale

            .PictureUnit2 = 5

        End With

    End If

End With
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

