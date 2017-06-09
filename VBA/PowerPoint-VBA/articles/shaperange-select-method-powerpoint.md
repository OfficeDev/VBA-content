---
title: ShapeRange.Select Method (PowerPoint)
keywords: vbapp10.chm548052
f1_keywords:
- vbapp10.chm548052
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Select
ms.assetid: 475f035e-a266-c263-eb62-542c51bb4087
ms.date: 06/08/2017
---


# ShapeRange.Select Method (PowerPoint)

Selects the specified object.


## Syntax

 _expression_. **Select**( **_Replace_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Replace_|Optional|**MsoTriState**|Specifies whether the selection replaces any previous selection.|

## Remarks

If you try to make a selection that isn't appropriate for the view, your code will fail. For example, you can select a slide in slide sorter view but not in slide view.

The  _Replace_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The selection is added to the previous selection.|
|**msoTrue**|The default. The selection replaces any previous selection.|

## Example

This example selects shapes one and three on slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Shapes.Range(Array(1, 3)).Select
```

This example adds shapes two and four on slide one in the active presentation to the previous selection.




```vb
ActivePresentation.Slides(1).Shapes.Range(Array(2, 4)).Select
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

