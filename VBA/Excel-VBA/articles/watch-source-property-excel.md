---
title: Watch.Source Property (Excel)
keywords: vbaxl10.chm690074
f1_keywords:
- vbaxl10.chm690074
ms.prod: excel
api_name:
- Excel.Watch.Source
ms.assetid: d21d19fb-cef2-b1c9-b3b7-4393ccbcec8c
ms.date: 06/08/2017
---


# Watch.Source Property (Excel)

Returns a  **Variant** value that represents the unique name that identifies items that have a **[SourceType](listobject-sourcetype-property-excel.md)** property value of **xlSourceRange** , **xlSourceChart** , **xlSourcePrintArea** , **xlSourceAutoFilter** , **xlSourcePivotTable** , or **xlSourceQuery** .


## Syntax

 _expression_ . **Source**

 _expression_ A variable that represents a **[Watch](watch-object-excel.md)** object.


## Remarks

If the  **SourceType** property is set to **xlSourceRange** , this property returns a range, which can be a defined name. If the **SourceType** property is set to **xlSourceChart** , **xlSourcePivotTable** , or **xlSourceQuery** , this property returns the name of the object, such as a chart name, a PivotTable report name, or a query table name.


## See also


#### Concepts


[Watch Object](watch-object-excel.md)

