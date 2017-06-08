---
title: CalculatedMember.DisplayFolder Property (Excel)
keywords: vbaxl10.chm686082
f1_keywords:
- vbaxl10.chm686082
ms.prod: excel
api_name:
- Excel.CalculatedMember.DisplayFolder
ms.assetid: 9ece45d1-4d27-0305-1189-15c414353607
ms.date: 06/08/2017
---


# CalculatedMember.DisplayFolder Property (Excel)

Returns the display folder name for a named set. Read-only


## Syntax

 _expression_ . **DisplayFolder**

 _expression_ A variable that returns a **[CalculatedMember](calculatedmember-object-excel.md)** object.


### Return Value

 **String**


## Remarks

The value of this property corresponds to the optional value that can be entered in the  **Display folder** text box of the **New/Modify Set** dialog box when a named set is created or edited. To create a new named set from data in a PivotTable based on an OLAP data source, click the PivotTable, click **Field, Items, &; Sets** on the **PivotTable Tools Options** tab on the ribbon, click **Manage Sets**, click  **New** in the ** Set Manager** dialog box, and then click **Create Set using MDX**. This will display the  **New Set** dialog box, which contains the **Display folder** text box. Similarly, if you select an existing named set in the **Set Manager** dialog box, and then click **Edit**, the  **Modify Set** dialog box is displayed.

This property along with the  **[Dynamic](calculatedmember-dynamic-property-excel.md)** and **[HierarchizeDistinct](calculatedmember-hierarchizedistinct-property-excel.md)** properties can only be read for named sets (which are represented by **[CalculatedMember](calculatedmember-object-excel.md)** objects where the **[Type](calculatedmember-type-property-excel.md)** property equals **xlCalculatedSet** ). These properties for cannot be read for calculated members or measures (which are represented by **CalculatedMember** objects where the **Type** property equals **xlCalculatedMember** ). If you attempt to read these properties for calculated members or measures, a run-time error is raised.


## See also


#### Concepts


[CalculatedMember Object](calculatedmember-object-excel.md)

