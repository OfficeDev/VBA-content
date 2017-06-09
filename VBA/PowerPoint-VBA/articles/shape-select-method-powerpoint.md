---
title: Shape.Select Method (PowerPoint)
keywords: vbapp10.chm547052
f1_keywords:
- vbapp10.chm547052
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Select
ms.assetid: 9fcf0ba4-ee6e-ecca-7948-7542db03ee99
ms.date: 06/08/2017
---


# Shape.Select Method (PowerPoint)

Selects the specified object.


## Syntax

 _expression_. **Select**( **_Replace_** )

 _expression_ A variable that represents a **Shape** object.


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

This example selects shape one on slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Shapes(1).Select
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

