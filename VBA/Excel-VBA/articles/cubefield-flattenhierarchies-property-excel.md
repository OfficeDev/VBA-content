---
title: CubeField.FlattenHierarchies Property (Excel)
keywords: vbaxl10.chm668103
f1_keywords:
- vbaxl10.chm668103
ms.prod: excel
api_name:
- Excel.CubeField.FlattenHierarchies
ms.assetid: bb97acc3-199b-6c40-e5b5-d411eb40b7e6
ms.date: 06/08/2017
---


# CubeField.FlattenHierarchies Property (Excel)

Returns or sets whether items from all levels of hierarchies in a named set cube field are displayed in the same field of a PivotTable report based on an OLAP cube. Read/write


## Syntax

 _expression_ . **FlattenHierarchies**

 _expression_ A variable that represents a **[CubeField](cubefield-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

 **True** if all hierarchies of the specified named set are displayed in the same field; otherwise **False** .

The value of this property corresponds to the setting of the  **Display items from different levels in separate fields** check box on the **Layout &; Print** tab of the **Field Settings** dialog box for a named set in a PivotTable report that is based on an OLAP cube.

This property returns an error if the  **[CubeFieldType](cubefield-cubefieldtype-property-excel.md)** property of the specified **CubeField** object is not **xlSet** .


## Example

The following code example flattens the hierarchies of the specified cube field, so that all levels are displayed in the same field of the PivotTable.


```vb
ActiveSheet.PivotTables("PivotTable1").CubeFields("[Summary P&;L]"). _ 
 FlattenHierarchies = True
```


## See also


#### Concepts


[CubeField Object](cubefield-object-excel.md)

