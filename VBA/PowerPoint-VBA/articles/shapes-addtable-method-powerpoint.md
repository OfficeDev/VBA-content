---
title: Shapes.AddTable Method (PowerPoint)
keywords: vbapp10.chm543027
f1_keywords:
- vbapp10.chm543027
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddTable
ms.assetid: 77ce193e-10f7-25f4-a6f8-99d7d2b781ad
ms.date: 06/08/2017
---


# Shapes.AddTable Method (PowerPoint)

Adds a table shape to a slide.


## Syntax

 _expression_. **AddTable**( **_NumRows_**, **_NumColumns_**, **_Left_**, **_Top_**, **_Width_**, **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumRows_|Required|**Long**|The number of rows in the table.|
| _NumColumns_|Required|**Long**|The number of columns in the table.|
| _Left_|Optional|**Single**|The distance (in points) from the left edge of the slide to the left edge of the table.|
| _Top_|Optional|**Single**|The distance (in points) from the top edge of the slide to the top edge of the table.|
| _Width_|Optional|**Single**|The width (in points) of the new table.|
| _Height_|Optional|**Single**|The height (in points) of the new table.|

### Return Value

Shape


## Example

This example creates a new table on slide two of the active presentation. The table has three rows and four columns. It is 10 points from the left edge of the slide, and 10 points from the top edge. The width of the new table is 288 points, which makes each of the four columns one inch wide (there are 72 points per inch). The height is set to 216 points, which makes each of the three rows one inch tall.


```vb
ActivePresentation.Slides(2).Shapes.AddTable(3, 4, 10, 10, 288, 216)
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

