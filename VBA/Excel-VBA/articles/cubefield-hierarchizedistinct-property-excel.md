---
title: CubeField.HierarchizeDistinct Property (Excel)
keywords: vbaxl10.chm668104
f1_keywords:
- vbaxl10.chm668104
ms.prod: excel
api_name:
- Excel.CubeField.HierarchizeDistinct
ms.assetid: 714f85b7-2adb-0ec1-5203-ca797b21e0a8
ms.date: 06/08/2017
---


# CubeField.HierarchizeDistinct Property (Excel)

Returns or sets whether to order and remove duplicates when displaying the specified named set in a PivotTable report based on an OLAP cube. Read/write


## Syntax

 _expression_ . **HierarchizeDistinct**

 _expression_ A variable that represents a **[CubeField](cubefield-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

 **True** if the named set is displayed as ordered with duplicates removed; otherwise **False** .

The value of this property corresponds to the setting of the  **Automatically order and remove duplicates from the set** check box on the **Layout &; Print** tab of the **Field Settings** dialog box for a named set in a PivotTable report based on an OLAP cube.

This property returns an error if the  **[CubeFieldType](cubefield-cubefieldtype-property-excel.md)** property of the specified **CubeField** object is not **xlSet** .


## Example

The following code example sets the  **HierarchizeDistinct** property to **True** to order and remove duplicates from the specified named set.


```vb
ActiveSheet.PivotTables("PivotTable1").CubeFields("[Summary P&;L]"). _ 
 HierarchizeDistinct = True
```


## See also


#### Concepts


[CubeField Object](cubefield-object-excel.md)

