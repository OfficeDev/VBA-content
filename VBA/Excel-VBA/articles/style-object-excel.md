---
title: Style Object (Excel)
keywords: vbaxl10.chm176072
f1_keywords:
- vbaxl10.chm176072
ms.prod: excel
api_name:
- Excel.Style
ms.assetid: 3c1e9184-0075-5f46-9a1a-0b61d874d1f8
ms.date: 06/08/2017
---


# Style Object (Excel)

Represents a style description for a range.


## Remarks

 The **Style** object contains all style attributes (font, number format, alignment, and so on) as properties. There are several built-in styles, including Normal, Currency, and Percent. Using the **Style** object is a fast and efficient way to change several cell-formatting properties on multiple cells at the same time.

For the  **[Workbook](workbook-object-excel.md)** object, the **Style** object is a member of the **[Styles](styles-object-excel.md)** collection. The **Styles** collection contains all the defined styles for the workbook.

You can change the appearance of a cell by changing properties of the style applied to that cell. Keep in mind, however, that changing a style property will affect all cells already formatted with that style.

Styles are sorted alphabetically by style name. The style index number denotes the position of the specified style in the sorted list of style names.  `Styles(1)` is the first style in the alphabetic list, and `Styles(Styles.Count)` is the last one in the list.

For more information about creating and modifying a style, see the  **[Styles](styles-object-excel.md)** object.


## Example

Use the  **Style** property to return the **Style** object used with a **Range** object. The following example applies the Percent style to cells A1:A10 on Sheet1.


```
Worksheets("Sheet1").Range("A1:A10").Style = "Percent"
```

Use  **Styles** ( _index_ ), where _index_ is the style index number or name, to return a single **Style** object from the workbook **Styles** collection. The following example changes the Normal style for the active workbook by setting the style's **Bold** property.




```
ActiveWorkbook.Styles("Normal").Font.Bold = True
```


## Methods



|**Name**|
|:-----|
|[Delete](style-delete-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[AddIndent](style-addindent-property-excel.md)|
|[Application](style-application-property-excel.md)|
|[Borders](style-borders-property-excel.md)|
|[BuiltIn](style-builtin-property-excel.md)|
|[Creator](style-creator-property-excel.md)|
|[Font](style-font-property-excel.md)|
|[FormulaHidden](style-formulahidden-property-excel.md)|
|[HorizontalAlignment](style-horizontalalignment-property-excel.md)|
|[IncludeAlignment](style-includealignment-property-excel.md)|
|[IncludeBorder](style-includeborder-property-excel.md)|
|[IncludeFont](style-includefont-property-excel.md)|
|[IncludeNumber](style-includenumber-property-excel.md)|
|[IncludePatterns](style-includepatterns-property-excel.md)|
|[IncludeProtection](style-includeprotection-property-excel.md)|
|[IndentLevel](style-indentlevel-property-excel.md)|
|[Interior](style-interior-property-excel.md)|
|[Locked](style-locked-property-excel.md)|
|[MergeCells](style-mergecells-property-excel.md)|
|[Name](style-name-property-excel.md)|
|[NameLocal](style-namelocal-property-excel.md)|
|[NumberFormat](style-numberformat-property-excel.md)|
|[NumberFormatLocal](style-numberformatlocal-property-excel.md)|
|[Orientation](style-orientation-property-excel.md)|
|[Parent](style-parent-property-excel.md)|
|[ReadingOrder](style-readingorder-property-excel.md)|
|[ShrinkToFit](style-shrinktofit-property-excel.md)|
|[Value](style-value-property-excel.md)|
|[VerticalAlignment](style-verticalalignment-property-excel.md)|
|[WrapText](style-wraptext-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
