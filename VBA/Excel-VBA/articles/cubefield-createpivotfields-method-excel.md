---
title: CubeField.CreatePivotFields Method (Excel)
keywords: vbaxl10.chm668099
f1_keywords:
- vbaxl10.chm668099
ms.prod: excel
api_name:
- Excel.CubeField.CreatePivotFields
ms.assetid: 87d868d7-8836-5a0b-a4b6-1ca3165b96e0
ms.date: 06/08/2017
---


# CubeField.CreatePivotFields Method (Excel)

 The **CreatePivotFields** method enables users to apply a filter to PivotFields not yet added to the PivotTable by creating the corresponding **PivotField** object.


## Syntax

 _expression_ . **CreatePivotFields**

 _expression_ A variable that represents a **CubeField** object.


## Remarks

In OLAP PivotTables, PivotFields do not exist until the corresponding CubeField is added to the PivotTable. The  **CreatePivotFields** method enables users to create all PivotFields of a CubeField. Users can also add filters to the PivotFields and set properties on them before the CubeField is added to the PivotTable.


## Example


```vb
Sub FilterFieldBeforeAddingItToPivotTable() 
 ActiveSheet.PivotTables("PivotTable1").CubeFields("[Date].[Fiscal]").CreatePivotFields 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields("[Date].[Fiscal].[Fiscal Year]").VisibleItemsList = 
 
 Array("[Date].[Fiscal].[Fiscal Year].&;[2003]", "[Date].[Fiscal].[Fiscal Year].&;[2004]", "[Date].[Fiscal].[Fiscal Year].&;[2005]") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields( _ 
 "[Date].[Fiscal].[Fiscal Semester]").VisibleItemsList = Array("") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields( _ 
 "[Date].[Fiscal].[Fiscal Quarter]").VisibleItemsList = Array("") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields("[Date].[Fiscal].[Month]"). _ 
 VisibleItemsList = Array("") 
 
 ActiveSheet.PivotTables("PivotTable1").PivotFields("[Date].[Fiscal].[Date]"). _ 
 VisibleItemsList = Array("") 
End Sub
```


## See also


#### Concepts


[CubeField Object](cubefield-object-excel.md)

