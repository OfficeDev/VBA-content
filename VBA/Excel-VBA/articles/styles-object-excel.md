---
title: Styles Object (Excel)
keywords: vbaxl10.chm178072
f1_keywords:
- vbaxl10.chm178072
ms.prod: excel
api_name:
- Excel.Styles
ms.assetid: 146effdc-e007-814d-b110-f7bd944fc15f
ms.date: 06/08/2017
---


# Styles Object (Excel)

A collection of all the  **[Style](style-object-excel.md)** objects in the specified or active workbook.


## Remarks

 Each **Style** object represents a style description for a range. The **Style** object contains all style attributes (font, number format, alignment, and so on) as properties. There are several built-in styles — including Normal, Currency, and Percent.


## Example

Use the  **[Styles](workbook-styles-property-excel.md)** property to return the **Styles** collection. The following example creates a list of style names on worksheet one in the active workbook.


```
For i = 1 To ActiveWorkbook.Styles.Count 
 Worksheets(1).Cells(i, 1) = ActiveWorkbook.Styles(i).Name 
Next
```

Use the  **[Add](styles-add-method-excel.md)** method to create a new style and add it to the collection. The following example creates a new style based on the Normal style, modifies the border and font, and then applies the new style to cells A25:A30.




```
With ActiveWorkbook.Styles.Add(Name:="Bookman Top Border") 
 .Borders(xlTop).LineStyle = xlDouble 
 .Font.Bold = True 
 .Font.Name = "Bookman" 
End With 
Worksheets(1).Range("A25:A30").Style = "Bookman Top Border"
```

Use  **Styles** ( _index_ ), where _index_ is the style index number or name, to return a single **Style** object from the workbook **Styles** collection. The following example changes the Normal style for the active workbook by setting its **Bold** property.




```
ActiveWorkbook.Styles("Normal").Font.Bold = True
```


## Methods



|**Name**|
|:-----|
|[Add](styles-add-method-excel.md)|
|[Merge](styles-merge-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](styles-application-property-excel.md)|
|[Count](styles-count-property-excel.md)|
|[Creator](styles-creator-property-excel.md)|
|[Item](styles-item-property-excel.md)|
|[Parent](styles-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
