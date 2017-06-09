---
title: Window.View Property (Excel)
keywords: vbaxl10.chm356127
f1_keywords:
- vbaxl10.chm356127
ms.prod: excel
api_name:
- Excel.Window.View
ms.assetid: 604ea4f4-8268-9939-cac3-2e082a2c4831
ms.date: 06/08/2017
---


# Window.View Property (Excel)

Returns or sets the view showing in the window. Read/write  **[XlWindowView](xlwindowview-enumeration-excel.md)** .


## Syntax

 _expression_ . **View**

 _expression_ A variable that represents a **Window** object.


## Remarks





| **XlWindowView** can be one of these **XlWindowView** constants.|
| **xlNormalView**|
| **xlPageBreakPreview**|
| **xlPageLayoutView**|

## Example

This example switches the view in the active window to page break preview.


```vb
ActiveWindow.View = xlPageBreakPreview
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

