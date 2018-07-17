---
title: IconSetCondition.ReverseOrder Property (Excel)
keywords: vbaxl10.chm812083
f1_keywords:
- vbaxl10.chm812083
ms.prod: excel
api_name:
- Excel.IconSetCondition.ReverseOrder
ms.assetid: cd42262e-06b0-04d5-c962-00f937d0d5dc
ms.date: 06/08/2017
---


# IconSetCondition.ReverseOrder Property (Excel)

Returns or sets a  **Boolean** value indicating if the order of icons is reversed for an icon set.


## Syntax

 _expression_ . **ReverseOrder**

 _expression_ A variable that represents an **[IconSetCondition](iconsetcondition-object-excel.md)** object.


## Remarks

By default, most of the icon sets you can use associate positive images with higher values. For example, the "3 Traffic Lights" icon set applies a green circle to the upper value threshold and a red circle to the lowest values in the range. If your data is such that lower values are more desirable, such as marathon time results, you may want to reverse the order that icons are applied to the range to associate the green circle to the lowest values.

If the  **[IconSet](iconsetcondition-iconset-property-excel.md)** property of the **IconSetCondition** object is **xlCustomSet** , the **ReverseOrder** property will return **False** . Additionally, if you attempt to set the **ReverseOrder** property to **True** when the **IconSet** property of the **IconSetCondition** object is **xlCustomSet** , Excel will return a runtime error.


## See also


#### Concepts


[IconSetCondition Object](iconsetcondition-object-excel.md)

