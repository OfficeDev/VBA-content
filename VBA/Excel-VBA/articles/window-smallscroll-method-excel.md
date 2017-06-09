---
title: Window.SmallScroll Method (Excel)
keywords: vbaxl10.chm356110
f1_keywords:
- vbaxl10.chm356110
ms.prod: excel
api_name:
- Excel.Window.SmallScroll
ms.assetid: dcf1fdeb-36ab-06ed-a9fc-5b2bbaecc457
ms.date: 06/08/2017
---


# Window.SmallScroll Method (Excel)

Scrolls the contents of the window by rows or columns.


## Syntax

 _expression_ . **SmallScroll**( **_Down_** , **_Up_** , **_ToRight_** , **_ToLeft_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Down_|Optional| **Variant**|The number of rows to scroll the contents down.|
| _Up_|Optional| **Variant**|The number of rows to scroll the contents up.|
| _ToRight_|Optional| **Variant**|The number of columns to scroll the contents to the right.|
| _ToLeft_|Optional| **Variant**|The number of columns to scroll the contents to the left.|

### Return Value

Variant


## Remarks

If  _Down_ and _Up_ are both specified, the contents of the window are scrolled by the difference of the arguments. For example, if _Down_ is 3 and _Up_ is 6, the contents are scrolled up three rows.

If  _ToLeft_ and _ToRight_ are both specified, the contents of the window are scrolled by the difference of the arguments. For example, if _ToLeft_ is 3 and _ToRight_ is 6, the contents are scrolled to the right three columns.

Any of these arguments can be a negative number.


## Example

This example scrolls the contents of the active window of Sheet1 down three rows.


```vb
Worksheets("Sheet1").Activate 
ActiveWindow.SmallScroll down:=3
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

