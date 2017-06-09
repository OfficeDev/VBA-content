---
title: PrintRanges.ClearAll Method (PowerPoint)
keywords: vbapp10.chm518003
f1_keywords:
- vbapp10.chm518003
ms.prod: powerpoint
api_name:
- PowerPoint.PrintRanges.ClearAll
ms.assetid: 3e177e7c-486e-a938-cf80-2f13b018094a
ms.date: 06/08/2017
---


# PrintRanges.ClearAll Method (PowerPoint)

Clears all the print ranges from the  **[PrintRanges](printranges-object-powerpoint.md)** collection. Use the **Add** method of the **PrintRanges** collection to add print ranges to the collection.


## Syntax

 _expression_. **ClearAll**

 _expression_ A variable that represents a **PrintRanges** object.


### Return Value

Nothing


## Example

This example clears any previously defined print ranges in the active presentation; creates new print ranges that contain slide 1, slides 3 through 5, and slides 8 and 9; prints the newly defined slide ranges; and then clears the new print ranges.


```vb
With ActivePresentation.PrintOptions

    .RangeType = ppPrintSlideRange

    With .Ranges

        .ClearAll

        .Add 1, 1

        .Add 3, 5

        .Add 8, 9

        .Parent.Parent.PrintOut

        .ClearAll

    End With

End With
```


## See also


#### Concepts


[PrintRanges Object](printranges-object-powerpoint.md)

