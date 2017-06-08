---
title: LinearGradient Object (Excel)
keywords: vbaxl10.chm854072
f1_keywords:
- vbaxl10.chm854072
ms.prod: excel
api_name:
- Excel.LinearGradient
ms.assetid: cb648564-0f57-f1b9-1c89-0329c110583f
ms.date: 06/08/2017
---


# LinearGradient Object (Excel)

The  **LinearGradient** object transitions through a series of colors in a linear manner along a specific angle.


## Remarks


- Attempting to access a Gradient property of an  **Interior** object that does not have an existing gradient fill will result in a Run Time Error. Be aware of the `Interior.Pattern` property before accessing the Gradient property.
    
- If [Interior.Pattern](interior-pattern-property-excel.md) is changed from a gradient type to a non-gradient type, the Gradient object will populate with default values.
    

 **Note**  Some things to consider when working with  **LinearGradient** objects


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

