---
title: PivotItem.SourceNameStandard Property (Excel)
keywords: vbaxl10.chm246093
f1_keywords:
- vbaxl10.chm246093
ms.prod: excel
api_name:
- Excel.PivotItem.SourceNameStandard
ms.assetid: f8e25ad0-7a97-c19c-85b5-bf25e3553ca8
ms.date: 06/08/2017
---


# PivotItem.SourceNameStandard Property (Excel)

Returns a  **String** that represents the PivotTable items' source name in standard English (United States) format settings. Read-only.


## Syntax

 _expression_ . **SourceNameStandard**

 _expression_ A variable that represents a **PivotItem** object.


## Remarks

This property is used when an item has a localized version and its  **SourceNameStandard** property value differs from the **[SourceName](pivotitem-sourcename-property-excel.md)** property value, such as with date formatting.


## Example

This example displays the source name for the sixth item on the fifth field of the active PivotTable. The example assumes that a PivotTable exists on the active worksheet and that the data source contains at least five fields and six items per field.


```vb
Sub CheckSourceNameStandard() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 Dim pvtItem As PivotItem 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields(5) 
 Set pvtItem = pvtField.PivotItems(6) 
 
 ' Display source name. 
 MsgBox "The source name is: " &; pvtItem.SourceNameStandard 
 
End Sub
```


## See also


#### Concepts


[SlicerItem Object](sliceritem-object-excel.md)
[PivotItem Object](pivotitem-object-excel.md)

