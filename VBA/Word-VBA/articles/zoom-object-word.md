---
title: Zoom Object (Word)
keywords: vbawd10.chm2470
f1_keywords:
- vbawd10.chm2470
ms.prod: word
api_name:
- Word.Zoom
ms.assetid: 9a07fe91-fe6c-21f8-7022-1c56676b89ef
ms.date: 06/08/2017
---


# Zoom Object (Word)

Contains magnification options (for example, the zoom percentage) for a window or pane. The  **Zoom** object is a member of the **[Zooms](zooms-object-word.md)** collection.


## Remarks

Use the  **Zoom** property of the **View** object to return a single **Zoom** object. The following example sets the zoom percentage for the active window to 110 percent.


```vb
ActiveDocument.ActiveWindow.View.Zoom.Percentage = 110
```

Use  **Zooms** (Index), where Index identifies the view type, to return a single **Zoom** object. The view type specified by index can be one of the following **[WdViewType](wdviewtype-enumeration-word.md)** constants: **wdMasterView** , **wdNormalView** , **wdOutlineView** , **wdPrintPreview** , **wdPrintView** , or **wdWebView** . The following example sets the magnification for the active window so that an entire page is visible.




```vb
ActiveDocument.ActiveWindow.ActivePane _ 
 .Zooms(wdPrintView).PageFit = wdPageFitFullPage
```

The  **Add** method isn't available for the **Zooms** collection. The **Zooms** collection includes a single **Zoom** object for each of the various view types (such as outline, normal, or page layout).


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


