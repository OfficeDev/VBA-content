---
title: QueryTable.BeforeRefresh Event (Excel)
keywords: vbaxl10.chm519073
f1_keywords:
- vbaxl10.chm519073
ms.prod: excel
api_name:
- Excel.QueryTable.BeforeRefresh
ms.assetid: 763cfe16-d48c-07f2-73e1-5c59021b4e58
ms.date: 06/08/2017
---


# QueryTable.BeforeRefresh Event (Excel)

Occurs before any refreshes of the query table. This includes refreshes resulting from calling the  **Refresh** method, from the user's actions in the product, and from opening the workbook containing the query table.


## Syntax

 _expression_ . **BeforeRefresh**( **_Cancel_** )

 _expression_ A variable that represents a **QueryTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the refresh doesn't occur when the procedure is finished.|

### Return Value

Nothing


## Example

This example runs before the query table is refreshed.


```vb
Private Sub QueryTable_BeforeRefresh(Cancel As Boolean) 
 a = MsgBox("Refresh Now?", vbYesNoCancel) 
 If a = vbNo Then Cancel = True 
 MsgBox Cancel 
End Sub
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

