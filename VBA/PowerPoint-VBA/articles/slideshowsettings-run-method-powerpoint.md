---
title: SlideShowSettings.Run Method (PowerPoint)
keywords: vbapp10.chm514008
f1_keywords:
- vbapp10.chm514008
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.Run
ms.assetid: 497fae3b-b6a3-dc26-20d9-bdc8057ddc09
ms.date: 06/08/2017
---


# SlideShowSettings.Run Method (PowerPoint)

Runs a slide show of the specified presentation. Returns a  **[SlideShowWindow](slideshowwindow-object-powerpoint.md)** object.


## Syntax

 _expression_. **Run**

 _expression_ A variable that represents a **SlideShowSettings** object.


### Return Value

SlideShowWindow


## Remarks

To run a custom slide show, set the  **RangeType** property to **ppShowNamedSlideShow**, and set the **SlideShowName** property to the name of the custom show you want to run.


## Example

This example starts a full-screen slide show of the active presentation, with shortcut keys disabled.


```vb
With ActivePresentation.SlideShowSettings

    .ShowType = ppShowSpeaker

    .Run.View.AcceleratorsEnabled = False

End With
```

This example runs the named slide show "Quick Show."




```vb
With ActivePresentation.SlideShowSettings

    .RangeType = ppShowNamedSlideShow

    .SlideShowName = "Quick Show"

    .Run

End With
```


## See also


#### Concepts


[SlideShowSettings Object](slideshowsettings-object-powerpoint.md)

