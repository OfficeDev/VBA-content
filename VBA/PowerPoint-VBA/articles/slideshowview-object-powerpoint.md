---
title: SlideShowView Object (PowerPoint)
keywords: vbapp10.chm513000
f1_keywords:
- vbapp10.chm513000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView
ms.assetid: 403b30ef-b12f-3a3c-e8d8-19189fd762fe
ms.date: 06/08/2017
---


# SlideShowView Object (PowerPoint)

Represents the view in a slide show window.


## Example

Use the [View](http://msdn.microsoft.com/library/ebf565af-fc90-ab1b-0e05-6dcb90a7c2d2%28Office.15%29.aspx)property of the  **SlideShowWindow** object to return the **SlideShowView** object. The following example sets slide show window one to display the first slide in the presentation.


```
SlideShowWindows(1).View.First
```

Use the [Run](http://msdn.microsoft.com/library/497fae3b-b6a3-dc26-20d9-bdc8057ddc09%28Office.15%29.aspx)method of the  **SlideShowSettings** object to create a **SlideShowWindow** object, and then use the **View** property to return the **SlideShowView** object the window contains. The following example runs a slide show of the active presentation, changes the pointer to a pen, and sets the pen color for the slide show to red.




```
With ActivePresentation.SlideShowSettings.Run.View

    .PointerColor.RGB = RGB(255, 0, 0)

    .PointerType = ppSlideShowPointerPen

End With
```


## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
[SlideShowView Object Members](http://msdn.microsoft.com/library/fe2aacef-7324-4d07-55e9-0dffcdbb2a6c%28Office.15%29.aspx)
