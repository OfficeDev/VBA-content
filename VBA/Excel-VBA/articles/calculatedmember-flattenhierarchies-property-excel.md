---
title: CalculatedMember.FlattenHierarchies Property (Excel)
keywords: vbaxl10.chm686084
f1_keywords:
- vbaxl10.chm686084
ms.prod: excel
api_name:
- Excel.CalculatedMember.FlattenHierarchies
ms.assetid: b0df471b-884a-fe43-b839-9de943720d0e
ms.date: 06/08/2017
---


# CalculatedMember.FlattenHierarchies Property (Excel)

Returns or sets whether items from all levels of the hierarchy of the specified named set are displayed in the same field of a PivotTable report based on an OLAP cube.  **Boolean** Read/write


## Syntax

 _expression_ . **FlattenHierarchies**

 _expression_ A variable that represents a **[CalculatedMember](calculatedmember-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

 **True** if items from all levels of the hierarchy of the named set are displayed in the same field; otherwise **False** .

The value of this property corresponds to the setting of the  **Display items from different levels in separate fields** check box in the **New/Modify Set** dialog box when a named set is created or edited. To create a new named set from data in a PivotTable based on an OLAP data source, click the PivotTable, click **Field, Items, &; Sets** on the **PivotTable Tools Options** tab on the ribbon, click **Manage Sets**, click  **New** in the ** Set Manager** dialog box, and then click **Create Set using MDX**. This will display the  **New Set** dialog box, which contains the **Display items from different levels in separate fields** check box. Similarly, if you select an existing named set in the **Set Manager** dialog box, and then click **Edit**, the  **Modify Set** dialog box is displayed.


## See also


#### Concepts


[CalculatedMember Object](calculatedmember-object-excel.md)

