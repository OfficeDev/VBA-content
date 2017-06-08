---
title: PrintOptions.PrintComments Property (PowerPoint)
keywords: vbapp10.chm517017
f1_keywords:
- vbapp10.chm517017
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.PrintComments
ms.assetid: 3c908a66-1db7-ef43-48a4-153f6095d041
ms.date: 06/08/2017
---


# PrintOptions.PrintComments Property (PowerPoint)

Sets or returns whether comments will be printed. Read/write.


## Syntax

 _expression_. **PrintComments**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

MsoTriState


## Remarks

The value of the  **PrintComments** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. Comments will not be printed.|
|**msoTrue**| Comments will be printed.|

## Example

This example instructs Microsoft PowerPoint to print comments.


```vb
Sub PrintPresentationComments

    ActivePresentation.PrintOptions.PrintComments = msoTrue

End Sub
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

