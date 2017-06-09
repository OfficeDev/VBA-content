---
title: Workbook.BeforeClose Event (Excel)
keywords: vbaxl10.chm503076
f1_keywords:
- vbaxl10.chm503076
ms.prod: excel
api_name:
- Excel.Workbook.BeforeClose
ms.assetid: 1c440637-8289-c6dd-24e0-1b2764fd1694
ms.date: 06/08/2017
---


# Workbook.BeforeClose Event (Excel)

Occurs before the workbook closes. If the workbook has been changed, this event occurs before the user is asked to save changes.


## Syntax

 _expression_ . **BeforeClose**( **_Cancel_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the close operation stops and the workbook is left open.|

### Return Value

Nothing


## Example

This example always saves the workbook if it's been changed.


```vb
Private Sub Workbook_BeforeClose(Cancel as Boolean) 
 If Me.Saved = False Then Me.Save 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

