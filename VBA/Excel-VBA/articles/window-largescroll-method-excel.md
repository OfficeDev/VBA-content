---
title: Window.LargeScroll Method (Excel)
keywords: vbaxl10.chm356097
f1_keywords:
- vbaxl10.chm356097
ms.prod: excel
api_name:
- Excel.Window.LargeScroll
ms.assetid: f3d74426-ece5-559f-c8c2-c356eb532217
ms.date: 06/08/2017
---


# Window.LargeScroll Method (Excel)

Scrolls the contents of the window by pages.


## Syntax

 _expression_ . **LargeScroll**( **_Down_** , **_Up_** , **_ToRight_** , **_ToLeft_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Down_|Optional| **Variant**|The number of pages to scroll the contents down.|
| _Up_|Optional| **Variant**|The number of pages to scroll the contents up.|
| _ToRight_|Optional| **Variant**|The number of pages to scroll the contents to the right.|
| _ToLeft_|Optional| **Variant**|The number of pages to scroll the contents to the left.|

### Return Value

Variant


## Remarks

If  _Down_ and _Up_ are both specified, the contents of the window are scrolled by the difference of the arguments. For example, if _Down_ is 3 and _Up_ is 6, the contents are scrolled up three pages.

If  _ToLeft_ and _ToRight_ are both specified, the contents of the window are scrolled by the difference of the arguments. For example, if _ToLeft_ is 3 and _ToRight_ is 6, the contents are scrolled to the right three pages.

Any of the arguments can be a negative number.


## Example

This example scrolls the contents of the active window of Sheet1 down three pages.


```vb
Worksheets("Sheet1").Activate 
ActiveWindow.LargeScroll down:=3
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

