---
title: CalculatedMember.ParentMember Property (Excel)
keywords: vbaxl10.chm686087
f1_keywords:
- vbaxl10.chm686087
ms.prod: excel
ms.assetid: 72711256-a4e4-0aa1-64d5-a4342a9ad4a6
ms.date: 06/08/2017
---


# CalculatedMember.ParentMember Property (Excel)

Returns the name of the parent member for the parent hierarchy.  **String** Read-only


## Syntax

 _expression_ . **ParentMember**

 _expression_ A variable that represents a[CalculatedMember](calculatedmember-object-excel.md) object.


## Remarks

The default parent member is determined by whatever has been defined by the cube designer or Analysis Services as the default member of the selected hierarchy. For example, if an "All" member exists for the selected hierarchy, then this is typically the default parent member on the cube.

If the selected parent hierarchy does not have an "All" parent member, another default parent member is defined, either by the cube designer or programmatically by Analysis Services. (If a default member is not specified by the cube designer, the Analysis Services engine automatically defines one.)


## Property value

 **STRING**


## See also


#### Concepts


[CalculatedMember Object](calculatedmember-object-excel.md)

