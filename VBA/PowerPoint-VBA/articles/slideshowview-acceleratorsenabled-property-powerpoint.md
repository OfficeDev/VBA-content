---
title: SlideShowView.AcceleratorsEnabled Property (PowerPoint)
keywords: vbapp10.chm513007
f1_keywords:
- vbapp10.chm513007
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.AcceleratorsEnabled
ms.assetid: 04db702f-af30-1868-0cab-17e692892e82
ms.date: 06/08/2017
---


# SlideShowView.AcceleratorsEnabled Property (PowerPoint)

Determines whether shortcut keys are enabled during a slide show. Read/write.


## Syntax

 _expression_. **AcceleratorsEnabled**

 _expression_ A variable that represents an **SlideShowView** object.


### Return Value

MsoTriState


## Remarks

If shortcut keys are disabled during a slide show, you can neither use the keyboard to navigate in the slide show nor press F1 to get a list of shortcut keys. You can still use the ESC key to exit the slide show.

The value of the  **AcceleratorsEnabled** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Shortcut keys are disabled during a slide show.|
|**msoTrue**| The default. Shortcut keys are enabled during a slide show.|

## Example

This example runs a slide show of the active presentation with shortcut keys disabled.


```vb
ActivePresentation.SlideShowSettings.Run _
    .View.AcceleratorsEnabled = False
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

