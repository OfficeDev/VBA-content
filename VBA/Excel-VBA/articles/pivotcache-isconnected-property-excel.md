---
title: PivotCache.IsConnected Property (Excel)
keywords: vbaxl10.chm227098
f1_keywords:
- vbaxl10.chm227098
ms.prod: excel
api_name:
- Excel.PivotCache.IsConnected
ms.assetid: 5c238338-c242-019c-1a29-08d2c87bc3be
ms.date: 06/08/2017
---


# PivotCache.IsConnected Property (Excel)

Returns  **True** if the **MaintainConnection** property is **True** and the PivotTable cache is currently connected to its source. Returns **False** if it is not currently connected to its source. Read-only **Boolean** .


## Syntax

 _expression_ . **IsConnected**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks

The  **IsConnected** property does not check to see if the connection is connected. Even if this property returns **True** , sending commands to the provider could result in an error if the connection is no longer valid.

Requires that the cache source type is external and that it is an OLE DB data source.


## Example

The following example determines if the cache is connected to its source and notifies the user. This example assumes that a PivotTable exists on the active worksheet.


```vb
Sub CheckIsConnected() 
 
 ' Handle run-time error if external source is not OLE DB. 
 On Error GoTo Not_OLEDB 
 
 ' Check connection setting and notify the user accordingly. 
 If Application.ActiveWorkbook.PivotCaches.Item(1).IsConnected = True Then 
 MsgBox "The PivotCache is currently connected to its source." 
 Else 
 MsgBox "The PivotCache is not currently connected to its source." 
 End If 
 Exit Sub 
 
Not_OLEDB: 
 MsgBox "The data source is not an OLE DB data source." 
 
End Sub
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

