---
title: SlideShowWindows Object (PowerPoint)
keywords: vbapp10.chm510000
f1_keywords:
- vbapp10.chm510000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindows
ms.assetid: aa4c7a38-32ea-c206-ce1f-d78094410f52
ms.date: 06/08/2017
---


# SlideShowWindows Object (PowerPoint)

A collection of all the  **[SlideShowWindow](slideshowwindow-object-powerpoint.md)** objects that represent the open slide shows in Microsoft PowerPoint.


## Example

Use the [SlideShowWindows](application-slideshowwindows-property-powerpoint.md)property to return the  **SlideShowWindows** collection. Use **SlideShowWindows** (index), where index is the window index number, to return a single **SlideShowWindow** object. The following example reduces the height of slide show window one if it is a full-screen window.


```vb
With SlideShowWindows(1)

    If .IsFullScreen Then

        .Height = .Height - 20

    End If

End With
```

Use the [Run](slideshowsettings-run-method-powerpoint.md)method to create a new slide show window and add it to the  **SlideShowWindows** collection. The following example runs a slide show of the active presentation.




```vb
ActivePresentation.SlideShowSettings.Run
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

