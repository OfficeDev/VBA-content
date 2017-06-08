---
title: PrintOptions.NumberOfCopies Property (PowerPoint)
keywords: vbapp10.chm517006
f1_keywords:
- vbapp10.chm517006
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.NumberOfCopies
ms.assetid: 6630ac4d-5c19-ad5f-f557-12e25e198e17
ms.date: 06/08/2017
---


# PrintOptions.NumberOfCopies Property (PowerPoint)

Returns or sets the number of copies of a presentation to be printed. Read/write.


## Syntax

 _expression_. **NumberOfCopies**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

Long


## Remarks

Specifying a value for the  **Copies** argument of the **[PrintOut](presentation-printout-method-powerpoint.md)** method sets the value of this property. The default value is 1.


## Example

This example prints three collated copies of the active presentation.


```vb
With ActivePresentation.PrintOptions

    .NumberOfCopies = 3

    .Collate = True

    .Parent.PrintOut

End With
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

