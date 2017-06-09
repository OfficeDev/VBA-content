---
title: Series.PictureUnit2 Property (Word)
keywords: vbawd10.chm123734617
f1_keywords:
- vbawd10.chm123734617
ms.prod: word
api_name:
- Word.Series.PictureUnit2
ms.assetid: 461c860f-ad4d-394a-508c-a53ef6b00bdb
ms.date: 06/08/2017
---


# Series.PictureUnit2 Property (Word)

Returns or sets the unit for each picture on the chart if the  **[PictureType](series-picturetype-property-word.md)** property is set to **xlStackScale** ; otherwise, this property is ignored. Read/write **Double** .


## Syntax

 _expression_ . **PictureUnit2**

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


## Example

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


[Series Object](series-object-word.md)

