---
title: PrintOptions.Ranges Property (PowerPoint)
keywords: vbapp10.chm517012
f1_keywords:
- vbapp10.chm517012
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.Ranges
ms.assetid: d0011261-a663-534d-204f-af2cd02f72be
ms.date: 06/08/2017
---


# PrintOptions.Ranges Property (PowerPoint)

Returns the  **[PrintRanges](printranges-object-powerpoint.md)** object, which represents the ranges of slides in the presentation to be printed. Read-only.


## Syntax

 _expression_. **Ranges**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

PrintRanges


## Remarks

If you don't want to print an entire presentation, you must use the  **[Add](printranges-add-method-powerpoint.md)** method to create a **[PrintRange](printrange-object-powerpoint.md)** object for each consecutive run of slides you want to print. For example, if you want to print slide 1, slides 3 through 5, and slides 8 and 9 in a specified presentation, you must create three **PrintRange** objects: one that represents slide 1; one that represents slides 3 through 5; and one that represents slides 8 and 9. For more information, see the example for this property.

The  **[RangeType](printoptions-rangetype-property-powerpoint.md)** property must be set to **ppPrintSlideRange** for the ranges in the **PrintRanges** collection to be applied.

To clear all the existing print ranges from the  **PrintRanges** collection, use the **[ClearAll](printranges-clearall-method-powerpoint.md)** method.

Specifying a value for the  **To** and **From** arguments of the **[PrintOut](presentation-printout-method-powerpoint.md)** method sets the contents of the **[PrintRanges](printranges-object-powerpoint.md)** object.


## Example

This example prints slide 1, slides 3 through 5, and slides 8 and 9 in the active presentation.


```vb
With ActivePresentation

    With .PrintOptions

        .RangeType = ppPrintSlideRange

        With .Ranges

            .Add 1, 1

            .Add 3, 5

            .Add 8, 9

        End With

    End With

    .PrintOut

End With
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

