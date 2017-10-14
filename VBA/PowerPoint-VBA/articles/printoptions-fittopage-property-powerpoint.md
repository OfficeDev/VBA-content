---
title: PrintOptions.FitToPage Property (PowerPoint)
keywords: vbapp10.chm517004
f1_keywords:
- vbapp10.chm517004
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions.FitToPage
ms.assetid: 53476904-fcbd-0a53-3e64-5c64799c8327
ms.date: 06/08/2017
---


# PrintOptions.FitToPage Property (PowerPoint)

Determines whether the slides will be scaled to fill the page they're printed on. Read/write.


## Syntax

 _expression_. **FitToPage**

 _expression_ A variable that represents a **PrintOptions** object.


### Return Value

MsoTriState


## Remarks

The value of the  **FitToPage** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. The slides will have the dimensions specified in the  **Page Setup** dialog box, whether or not those dimensions match the page they're printed on.|
|**msoTrue**| The specified slides will be scaled to fill the page they're printed on, regardless of the values in the **Height** and **Width** boxes in the **Page Setup** dialog box. (On the **Design** tab, click **Page Setup**.)|

## Example

This example prints the active presentation and scales each slide to fit the printed page.


```vb
With ActivePresentation

    .PrintOptions.FitToPage = msoTrue

    .PrintOut

End With


```


## See also


#### Concepts


[PrintOptions Object](printoptions-object-powerpoint.md)

