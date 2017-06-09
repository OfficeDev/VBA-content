---
title: ListColumn.XPath Property (Excel)
keywords: vbaxl10.chm738080
f1_keywords:
- vbaxl10.chm738080
ms.prod: excel
api_name:
- Excel.ListColumn.XPath
ms.assetid: a025d423-4291-c40a-96b9-e1f807d4b388
ms.date: 06/08/2017
---


# ListColumn.XPath Property (Excel)

Returns an  **[XPath](xpath-object-excel.md)** object that represents the Xpath of the element mapped to the specified **[Range](range-object-excel.md)** object. The context of the range determines whether or not the action succeeds or returns an empty object. Read-only.


## Syntax

 _expression_ . **XPath**

 _expression_ A variable that represents a **ListColumn** object.


## Remarks

The  **XPath** property is valid when the range it contains meets the following conditions:


- The range is a single cell.
    
- If the range consists of two or more cells, then one or the other must be true:
    
      1. If the cells contain XPath information, then all cells in the range must contain XPath information (that is, each cell is associated with one or more data maps), and all of the cells must have identical XPath content (that is, each cell contributes to the same set of data maps).
    
  2. All of the cells must contain no XPath information.
    
- The range does not contain discontinuous areas.
    
     **Note**  The header and totals row of a table are considered to contain XPath information.
Any ranges that don't meet the above conditions returns a runtime error.

If the range selection is valid, but none of the cells are mapped, Excel returns an  **XPath** object so that you can access the **SetValue** method to create a mapping.


## See also


#### Concepts


[ListColumn Object](listcolumn-object-excel.md)

