---
title: DataLabels.AutoText Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels.AutoText
ms.assetid: 6e964058-3cfa-ba02-b324-fc1e82beb3d3
ms.date: 06/08/2017
---


# DataLabels.AutoText Property (PowerPoint)

 **True** if all objects in the collection automatically generate appropriate text based on context. Read/write **Boolean**.


## Syntax

 _expression_. **AutoText**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-powerpoint.md)** object.


## Remarks

Setting the value of this property sets the  **[AutoText](datalabel-autotext-property-powerpoint.md)** property of all **[DataLabel](datalabel-object-powerpoint.md)** objects contained by the collection. This property returns **True** only when the **AutoText** property for all **DataLabel** objects contained in the collection is set to **True**; otherwise, this property returns **False**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the data labels for series one of the first chart in the active document to automatically generate appropriate text.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1). _
            DataLabels.AutoText = True
    End If
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-powerpoint.md)

