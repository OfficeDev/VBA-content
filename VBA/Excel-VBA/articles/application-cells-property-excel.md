---
title: Application.Cells Property (Excel)
keywords: vbaxl10.chm183085
f1_keywords:
- vbaxl10.chm183085
ms.prod: excel
api_name:
- Excel.Application.Cells
ms.assetid: 9788c893-13c3-eb57-bcf7-50806b476ba3
ms.date: 06/08/2017
---


# Application.Cells Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents all the cells on the active worksheet. If the active document is not a worksheet, this property fails.


## Syntax

 _expression_ . **Cells**

 _expression_ A variable that represents an **Application** object.


## Remarks

Because the  **[Item](range-item-property-excel.md)** property is the default property for the **Range** object, you can specify the row and column index immediately after the **Cells** keyword. For more information, see the **Item** property and the examples for this topic.

Using this property without an object qualifier returns a  **Range** object that represents all the cells on the active worksheet.


## Example

 **Sample code provided by:** Tom Urtis,[Atlas Programming Management](http://www.atlaspm.com/)

This example looks at data in each row and inserts a blank row each time the value in column A changes.




```vb
Sub ChangeInsertRows()
    Application.ScreenUpdating = False
    Dim xRow As Long
    
    For xRow = Application.Cells(Rows.Count, 1).End(xlUp).Row To 3 Step -1
        If Application.Cells(xRow, 1).Value <> Application.Cells(xRow - 1, 1).Value Then Rows(xRow).Resize(1).Insert
    Next xRow
    
    Application.ScreenUpdating = True
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the coauthor of "Holy Macro! It's 2,500 Excel VBA Examples." 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Application Object](application-object-excel.md)

