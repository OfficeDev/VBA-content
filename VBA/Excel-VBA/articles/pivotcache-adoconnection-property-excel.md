---
title: PivotCache.ADOConnection Property (Excel)
keywords: vbaxl10.chm227097
f1_keywords:
- vbaxl10.chm227097
ms.prod: excel
api_name:
- Excel.PivotCache.ADOConnection
ms.assetid: 410a3eee-0dda-4be1-45c4-809893de624e
ms.date: 06/08/2017
---


# PivotCache.ADOConnection Property (Excel)

Returns an  **ADO Connection** object if the PivotTable cache is connected to an OLE DB data source. The **ADOConnection** property exposes the Microsoft Excel connection to the data provider, allowing the user to write code within the context of the same session that Excel is using with ADO (relational source) or ADO MD (OLAP source). Read-only.


## Syntax

 _expression_ . **ADOConnection**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks

The  **ADOConnection** property is available only for sessions with an OLE DB data source. When there is no ADO session, the query will result in a run-time error.

The  **ADOConnection** property can be used for any OLE DB-based cache with ADO. The **ADO Connection** object can be used with ADO MD for finding information about OLAP cubes on which the cache is based.


## Example

This example sets an  **ADO DB Connection** object to the **ADOConnection** property of the PivotTable cache. The example assumes that a PivotTable report exists on the active worksheet.


```vb
Sub UseADOConnection() 
 
 Dim ptOne As PivotTable 
 Dim cmdOne As New ADODB.Command 
 Dim cfOne As CubeField 
 
 Set ptOne = Sheet1.PivotTables(1) 
 ptOne.PivotCache.MaintainConnection = True 
 Set cmdOne.ActiveConnection = ptOne.PivotCache.ADOConnection 
 
 ptOne.PivotCache.MakeConnection 
 
 ' Create a set. 
 cmdOne.CommandText = "Create Set [Warehouse].[My Set] as '{[Product].[All Products].Children}'" 
 cmdOne.CommandType = adCmdUnknown 
 cmdOne.Execute 
 
 ' Add a set to the CubeField. 
 Set cfOne = ptOne.CubeFields.AddSet("My Set", "My Set") 
 
End Sub
```

This example adds a calculated member, assuming that a PivotTable report exists on the active worksheet.




```vb
Sub AddMember() 
 
 Dim cmd As New ADODB.Command 
 
 If Not ActiveSheet.PivotTables(1).PivotCache.IsConnected Then 
 ActiveSheet.PivotTables(1).PivotCache.MakeConnection 
 End If 
 
 Set cmd.ActiveConnection = ActiveSheet.PivotTables(1).PivotCache.ADOConnection 
 
 ' Add a calculated member. 
 cmd.CommandText = "CREATE MEMBER [Warehouse].[Product].[All Products].[Drink and Non-Consumable] AS '[Product].[All Products].[Drink] + [Product].[All Products].[Non-Consumable]'" 
 cmd.CommandType = adCmdUnknown 
 cmd.Execute 
 
 ActiveSheet.PivotTables(1).PivotCache.Refresh 
 
End Sub
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

