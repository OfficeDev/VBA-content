---
title: Application.DataEntryMode Property (Excel)
keywords: vbaxl10.chm133102
f1_keywords:
- vbaxl10.chm133102
ms.prod: excel
api_name:
- Excel.Application.DataEntryMode
ms.assetid: 1fd9f191-3aa5-2548-2d41-b9d2bc79654b
ms.date: 06/08/2017
---


# Application.DataEntryMode Property (Excel)

Returns or sets Data Entry mode, as shown in the following table. When in Data Entry mode, you can enter data only in the unlocked cells in the currently selected range. Read/write  **Long** .


## Syntax

 _expression_ . **DataEntryMode**

 _expression_ A variable that represents an **Application** object.


## Remarks





|**Value**|**Meaning**|
|:-----|:-----|
| **xlOn**|Data Entry mode is turned on.|
| **xlOff**|Data Entry mode is turned off.|
| **xlStrict**|Data Entry mode is turned on, and pressing ESC won't turn it off.|

## Example

This example turns off Data Entry mode if it's on.


```vb
If (Application.DataEntryMode = xlOn) Or _ 
 (Application.DataEntryMode = xlStrict) Then 
 Application.DataEntryMode = xlOff 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

