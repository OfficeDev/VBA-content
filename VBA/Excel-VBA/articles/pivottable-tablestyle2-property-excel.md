---
title: PivotTable.TableStyle2 Property (Excel)
keywords: vbaxl10.chm235171
f1_keywords:
- vbaxl10.chm235171
ms.prod: excel
api_name:
- Excel.PivotTable.TableStyle2
ms.assetid: d2d79fc6-2ead-91a9-f304-92248584f4b2
ms.date: 06/08/2017
---


# PivotTable.TableStyle2 Property (Excel)

The  **TableStyle2** property specifies the PivotTable style currently applied to the PivotTable. Read/write.


## Syntax

 _expression_ . **TableStyle2**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

The property is called TableStyle2 because there is an exisiting property named  **TableStyle** .


## Example


```vb
Sub ApplyingStyle() 
 
 ActiveSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight17" 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

