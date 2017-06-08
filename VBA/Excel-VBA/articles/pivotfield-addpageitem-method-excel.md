---
title: PivotField.AddPageItem Method (Excel)
keywords: vbaxl10.chm240136
f1_keywords:
- vbaxl10.chm240136
ms.prod: excel
api_name:
- Excel.PivotField.AddPageItem
ms.assetid: c7f63c9f-9ad2-fcd9-13de-e9e46c40b8dc
ms.date: 06/08/2017
---


# PivotField.AddPageItem Method (Excel)

Adds an additional item to a multiple item page field.


## Syntax

 _expression_ . **AddPageItem**( **_Item_** , **_ClearList_** )

 _expression_ A variable that represents a **PivotField** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **String**| Source name of a **PivotItem** object, corresponding to the specific Online Analytical Processing (OLAP) member unique name.|
| _ClearList_|Optional| **Variant**|If  **False** (default), adds a page item to the existing list. If **True** , deletes all current items and adds _Item_.|

## Remarks

To avoid run-time errors, the data source must be an OLAP source, the field chosen must currently be in the page position, and the  **[EnableMultiplePageItems](pivotfield-enablemultiplepageitems-property-excel.md)** property must be set to **True** .


## Example

In this example, Microsoft Excel adds a page item with a source name titled "[Product].[All Products].[Food].[Eggs]". This example assumes an OLAP PivotTable exists on the active worksheet.


```vb
Sub UseAddPageItem() 
 
 ' The source is an OLAP database and you can manually reorder items. 
 ActiveSheet.PivotTables(1).CubeFields("[Product]"). _ 
 EnableMultiplePageItems = True 
 
 ' Add the page item titled "[Product].[All Products].[Food].[Eggs]". 
 ActiveSheet.PivotTables(1).PivotFields("[Product]").AddPageItem ( _ 
 "[Product].[All Products].[Food].[Eggs]") 
 
End Sub
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

