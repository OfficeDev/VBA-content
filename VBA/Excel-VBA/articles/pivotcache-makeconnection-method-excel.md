---
title: PivotCache.MakeConnection Method (Excel)
keywords: vbaxl10.chm227099
f1_keywords:
- vbaxl10.chm227099
ms.prod: excel
api_name:
- Excel.PivotCache.MakeConnection
ms.assetid: d0b29374-4d5a-7d9e-630a-500b505da1bd
ms.date: 06/08/2017
---


# PivotCache.MakeConnection Method (Excel)

Establishes a connection for the specified PivotTable cache.


## Syntax

 _expression_ . **MakeConnection**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks

The  **MakeConnection** method can be used after the cache drops a connection and the user wants to reestablish the connection.

Various objects and methods might return a run-time error if the cache is not connected. Use of this method assures a connection before executing other objects or methods.

This method will result in a run-time error if the  **MaintainConnection** property of the specified PivotTable cache has been set to **False** , the **SourceType** property of the specified PivotTable cache has not been set to xlExternal, or if the connection is not to an OLE DB data source.


 **Note**  Microsoft Excel might drop a connection temporarily in the course of a session (unknown to the VBA programmer), so this method proves useful.


## Example

The following example determines if the cache is connected to its source and makes a connection to the source if necessary. This example assumes that a PivotTable cache exists on the active worksheet.


```vb
Sub UseMakeConnection() 
 
    Dim pvtCache As PivotCache 
 
    Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
    ' Handle run-time error if external source is not an OLE DB data source. 
    On Error GoTo Not_OLEDB 
 
    ' Check connection setting and make connection if necessary. 
    If pvtCache.IsConnected = True Then 
        MsgBox "The MakeConnection method is not needed." 
    Else 
        pvtCache.MakeConnection 
        MsgBox "A connection has been made." 
    End If 
    Exit Sub 
 
Not_OLEDB: 
    MsgBox "The data source is not an OLE DB data source" 
 
End Sub
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

