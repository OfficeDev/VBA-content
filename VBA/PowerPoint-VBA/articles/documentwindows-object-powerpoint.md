---
title: DocumentWindows Object (PowerPoint)
keywords: vbapp10.chm509000
f1_keywords:
- vbapp10.chm509000
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindows
ms.assetid: 84ed4b8c-593a-8100-d4b8-158115c4e84d
ms.date: 06/08/2017
---


# DocumentWindows Object (PowerPoint)

A collection of all the  **[DocumentWindow](documentwindow-object-powerpoint.md)** objects that are currently open in Microsoft PowerPoint. This collection doesn't include open slide show windows, which are included in the **[SlideShowWindows](slideshowwindows-object-powerpoint.md)** collection.


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.DocumentWindows**
    

## Example

Use the [Windows](application-windows-property-powerpoint.md) property to return the **DocumentWindows** collection. The following example tiles the open document windows.


```
Windows.Arrange ppArrangeTiled
```

Use the  **[NewWindow](documentwindow-newwindow-method-powerpoint.md)** method to create a document window and add it to the **DocumentWindows** collection. The following example creates a new window for the active presentation.




```vb
ActivePresentation.NewWindow
```

Use  **Windows** (index), where index is the window index number, to return a single **DocumentWindow** object. The following example closes document window two.




```
Windows(2).Close
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)


