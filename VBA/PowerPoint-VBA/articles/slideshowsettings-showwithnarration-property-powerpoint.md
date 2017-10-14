---
title: SlideShowSettings.ShowWithNarration Property (PowerPoint)
keywords: vbapp10.chm514011
f1_keywords:
- vbapp10.chm514011
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.ShowWithNarration
ms.assetid: 65390c53-abeb-ca9e-0697-f68dcb455324
ms.date: 06/08/2017
---


# SlideShowSettings.ShowWithNarration Property (PowerPoint)

Determines whether the specified slide show is shown with narration. Read/write.


## Syntax

 _expression_. **ShowWithNarration**

 _expression_ A variable that represents a **SlideShowSettings** object.


### Return Value

MsoTriState


## Remarks

The value of the  **ShowWithNarration** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified slide show is shown without narration. |
|**msoTrue**| The specified slide show is shown with narration.|

## Example

This example runs a slide show of the active presentation with animation and narration turned off.


```vb
With ActivePresentation.SlideShowSettings

    .ShowWithAnimation = msoFalse

    .ShowWithNarration = msoFalse

    .Run

End With
```


## See also


#### Concepts


[SlideShowSettings Object](slideshowsettings-object-powerpoint.md)

