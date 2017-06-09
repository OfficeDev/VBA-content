---
title: PrintOptions.Collate Property (PowerPoint)
keywords: vbapp10.chm517003
f1_keywords:
- vbapp10.chm517003
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.Collate
ms.assetid: 4cf1d714-6ea2-fce5-340e-202d91ad1137
ms.date: 06/08/2017
---


# PrintOptions.Collate Property (PowerPoint)

Determines whether a complete copy of the specified presentation is printed before the first page of the next copy is printed. Read/write.


## Syntax

 _expression_. **Collate**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

MsoTriState


## Remarks

Specifying a value for the  **Collate** argument of the **[PrintOut](presentation-printout-method-powerpoint.md)** method sets the value of this property.

The value of the  **Collate** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| A copy of the specified presentation is not printed before the first page of the next copy is printed.|
|**msoTrue**| The default. A complete copy of the specified presentation is printed before the first page of the next copy is printed.|

## Example

This example prints three collated copies of the active presentation.


```vb
With ActivePresentation.PrintOptions

    .NumberOfCopies = 3

    .Collate = msoTrue

    .Parent.PrintOut

End With
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

