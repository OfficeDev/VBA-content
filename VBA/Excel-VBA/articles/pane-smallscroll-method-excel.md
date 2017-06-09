---
title: Pane.SmallScroll Method (Excel)
keywords: vbaxl10.chm360078
f1_keywords:
- vbaxl10.chm360078
ms.prod: excel
api_name:
- Excel.Pane.SmallScroll
ms.assetid: d41345f6-1b46-0772-afba-81d377acc90f
ms.date: 06/08/2017
---


# Pane.SmallScroll Method (Excel)

Scrolls the contents of the window by rows or columns.


## Syntax

 _expression_ . **SmallScroll**( **_Down_** , **_Up_** , **_ToRight_** , **_ToLeft_** )

 _expression_ A variable that represents a **Pane** object.


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


## See also


#### Concepts


[Pane Object](pane-object-excel.md)

