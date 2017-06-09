---
title: PivotTable.VisualTotalsForSets Property (Excel)
keywords: vbaxl10.chm235200
f1_keywords:
- vbaxl10.chm235200
ms.prod: excel
api_name:
- Excel.PivotTable.VisualTotalsForSets
ms.assetid: c4a01954-ab23-433b-1e82-8450e752251f
ms.date: 06/08/2017
---


# PivotTable.VisualTotalsForSets Property (Excel)

Returns or sets whether to include filtered items in the totals of named sets for the specified PivotTable. Read/write


## Syntax

 _expression_ . **VisualTotalsForSets**

 _expression_ A variable that represents a **[PivotTable](pivottable-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

 **True** if filtered items are included in the totals for named sets; otherwise **False** . The default value of this property is **False** .

In a PivotTable based on an OLAP data source, you can configure whether totals for named sets in the PivotTable will include items that have been filtered. The setting of the  **VisualTotalsForSets** property corresponds to the **Include filtered items in set totals** check box on the **Totals &; Filters** tab of the **PivotTable Options** dialog box.


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

