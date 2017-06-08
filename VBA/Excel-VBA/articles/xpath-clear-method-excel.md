---
title: XPath.Clear Method (Excel)
keywords: vbaxl10.chm760077
f1_keywords:
- vbaxl10.chm760077
ms.prod: excel
api_name:
- Excel.XPath.Clear
ms.assetid: 8d9e0c70-c77e-257f-6ac7-7a8577282ab1
ms.date: 06/08/2017
---


# XPath.Clear Method (Excel)

Clears all XPath schema information for the mapped range. 


## Syntax

 _expression_ . **Clear**

 _expression_ A variable that represents a **XPath** object.


## Remarks

 **Clear** affects the entire range mapped to this **[XPath](xpath-object-excel.md)** object.

This method does not clear the data from the cells mapped to the specified XPath. Use the  **[Clear](range-clear-method-excel.md)** method of the **[Range](range-object-excel.md)** object to clear the data from the cells.

If the specified XPath is mapped in an XML list, then the schema mapping is removed, but the list is not deleted from the worksheet.

If the mapped range is a single-cell then the single-cell is removed and the data remains.

This method will produce an error if any of the following conditions are true:


- The range spans multiple columns in the grid.
    
- Part of the range spans already mapped cells and the rest spans unmapped cells.
    
- Part of the range spans one mapping, and another part of the range spans a different mapping or different XPath from the same mapping.
    

## See also


#### Concepts


[XPath Object](xpath-object-excel.md)

