---
title: Pane.LargeScroll Method (Excel)
keywords: vbaxl10.chm360075
f1_keywords:
- vbaxl10.chm360075
ms.prod: excel
api_name:
- Excel.Pane.LargeScroll
ms.assetid: e785bf52-d19b-a6e6-212b-0c5b5de88910
ms.date: 06/08/2017
---


# Pane.LargeScroll Method (Excel)

Scrolls the contents of the window by pages.


## Syntax

 _expression_ . **LargeScroll**( **_Down_** , **_Up_** , **_ToRight_** , **_ToLeft_** )

 _expression_ A variable that represents a **Pane** object.


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


## See also


#### Concepts


[Pane Object](pane-object-excel.md)

