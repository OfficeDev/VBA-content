---
title: CalculatedMember.Dynamic Property (Excel)
keywords: vbaxl10.chm686081
f1_keywords:
- vbaxl10.chm686081
ms.prod: excel
api_name:
- Excel.CalculatedMember.Dynamic
ms.assetid: b201fe58-1320-1fe0-8045-ab17b7543eee
ms.date: 06/08/2017
---


# CalculatedMember.Dynamic Property (Excel)

Returns whether the specified named set is recalculated with every update. Read-only


## Syntax

 _expression_ . **Dynamic**

 _expression_ A variable that returns a **[CalculatedMember](calculatedmember-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

 **True** if the named set is recalculated with every update; otherwise **False**.

The value of this property corresponds to the setting of the  **Recalculate set with every update** check box in the **New/Modify Set** dialog box that is available when a named set is created or edited. To create a new named set from data in a PivotTable based on an OLAP data source, click the PivotTable, click **Field, Items, &; Sets** on the **PivotTable Tools Options** tab on the ribbon, click **Manage Sets**, click  **New** in the ** Set Manager** dialog box, and then click **Create Set using MDX**. This will display the  **New Set** dialog box, which contains the **Recalculate set with every update** check box. Similarly, if you select an existing named set in the **Set Manager** dialog box, and then click **Edit**, the  **Modify Set** dialog box is displayed.



This property along with the  **[DisplayFolder](calculatedmember-displayfolder-property-excel.md)** and **[HierarchizeDistinct](calculatedmember-hierarchizedistinct-property-excel.md)** properties can only be read for named sets (which are represented by **[CalculatedMember](calculatedmember-object-excel.md)** objects where the **[Type](calculatedmember-type-property-excel.md)** property equals **xlCalculatedSet** ). These properties for cannot be read for calculated members or measures (which are represented by **CalculatedMember** objects where the **Type** property equals **xlCalculatedMember** ). If you attempt to read these properties for calculated members or measures, a run-time error is raised.


## See also


#### Concepts


[CalculatedMember Object](calculatedmember-object-excel.md)

