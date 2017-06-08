---
title: PrintOptions.PrintInBackground Property (PowerPoint)
keywords: vbapp10.chm517010
f1_keywords:
- vbapp10.chm517010
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.PrintInBackground
ms.assetid: d3a734a3-fa17-2321-1c29-6167f0110bd7
ms.date: 06/08/2017
---


# PrintOptions.PrintInBackground Property (PowerPoint)

Determines whether the specified presentation is printed in the background. Read/write.


## Syntax

 _expression_. **PrintInBackground**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

MsoTriState


## Remarks

The value of the  **PrintInBackground** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified presentation is not printed in the background.|
|**msoTrue**| The default. The specified presentation is printed in the background, which means that you can continue to work while it is being printed.|

## Example

This example prints the active presentation in the background.


```vb
With ActivePresentation

    .PrintOptions.PrintInBackground = msoTrue

    .PrintOut

End With
```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

