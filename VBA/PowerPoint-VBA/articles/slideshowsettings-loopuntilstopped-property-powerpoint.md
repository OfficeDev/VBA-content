---
title: SlideShowSettings.LoopUntilStopped Property (PowerPoint)
keywords: vbapp10.chm514009
f1_keywords:
- vbapp10.chm514009
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.LoopUntilStopped
ms.assetid: 767a5865-b50b-d7c6-6076-6786b43c6b88
ms.date: 06/08/2017
---


# SlideShowSettings.LoopUntilStopped Property (PowerPoint)

Determines whether specified slide show loops continuously until the user presses ESC. Read/write.


## Syntax

 _expression_. **LoopUntilStopped**

 _expression_ A variable that represents a **SlideShowSettings** object.


## Remarks

The value of the  **LoopUntilStopped** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified slide show does not loop continuously. |
|**msoTrue**| The specified slide show loops continuously until the user presses ESC.|

## Example

This example starts a slide show of the active presentation that will automatically advance the slides (using the stored timings) and will loop continuously through the presentation until the user presses ESC.


```vb
With ActivePresentation.SlideShowSettings

    .AdvanceMode = ppSlideShowUseSlideTimings

    .LoopUntilStopped = msoTrue

    .Run

End With
```


## See also


#### Concepts


[SlideShowSettings Object](slideshowsettings-object-powerpoint.md)

