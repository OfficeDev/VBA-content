---
title: VPageBreaks Object (Excel)
keywords: vbaxl10.chm166072
f1_keywords:
- vbaxl10.chm166072
ms.prod: excel
api_name:
- Excel.VPageBreaks
ms.assetid: ab8f288a-5235-76c9-7b27-81e542cdd141
ms.date: 06/08/2017
---


# VPageBreaks Object (Excel)

A collection of vertical page breaks within the print area.


## Remarks

Each vertical page break is represented by a  **[VPageBreak](vpagebreak-object-excel.md)** object.

When the [Application](vpagebreaks-application-property-excel.md) property, **[Count](vpagebreaks-count-property-excel.md)** property, **[Creator](lineformat-creator-property-excel.md)** property, **[Item](vpagebreaks-item-property-excel.md)** property, **[Parent](vpagebreaks-parent-property-excel.md)** property or **[Add](vpagebreaks-add-method-excel.md)** method is used in conjunction with the **VPageBreaks** property:


- For an automatic print area, the  **VPageBreaks** property applies only to the page breaks within the print area.
    
- For a user-defined print area of the same range, the  **VPageBreaks** property applies to all of the page breaks.
    

## Example

Use the  **[VPageBreaks](sheets-vpagebreaks-property-excel.md)** property to return the **VPageBreaks** collection. Use the **[Add](vpagebreaks-add-method-excel.md)** method to add a vertical page break.

If you add a page break that does not intersect the print area, then the newly-added  **VPageBreak** object will not appear in the **VPageBreaks** collection for the print area. The contents of the collection may change if the print area is resized or redefined.

The following example adds a vertical page break to the left of the active cell.




```vb
ActiveSheet.VPageBreaks.Add Before:=ActiveCell
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


