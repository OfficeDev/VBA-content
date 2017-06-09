---
title: Name Object (Excel)
keywords: vbaxl10.chm489072
f1_keywords:
- vbaxl10.chm489072
ms.prod: excel
api_name:
- Excel.Name
ms.assetid: cfedb297-ac0d-dff0-99c7-6927cc5f31ed
ms.date: 06/08/2017
---


# Name Object (Excel)

Represents a defined name for a range of cells. Names can be either built-in names — such as Database, Print_Area, and Auto_Open — or custom names.


## Remarks

 **Application, Workbook, and Worksheet Objects**

The  **Name** object is a member of the **[Names](names-object-excel.md)** collection for the **[Application](application-object-excel.md)**, **[Workbook](workbook-object-excel.md)**, and **[Worksheet](worksheet-object-excel.md)** objects. Use **[Names](workbook-names-property-excel.md)** ( _index_ ), where _index_ is the name index number or defined name, to return a single **Name** object.

The index number indicates the position of the name within the collection. Names are placed in alphabetic order, from a to z, and are not case-sensitive.

 **Range Objects**

Although a  **[Range](range-object-excel.md)** object can have more than one name, there's no **Names** collection for the **Range** object. Use **[Name](range-name-property-excel.md)** with a **Range** object to return the first name from the list of names (sorted alphabetically) assigned to the range. The following example sets the **[Visible](worksheet-visible-property-excel.md)** property for the first name assigned to cells A1:B1 on worksheet one.


## Example

The following example displays the cell reference for the first name in the application collection.


```
MsgBox Names(1).RefersTo
```

The following example deletes the name "mySortRange" from the active workbook.




```
ActiveWorkbook.Names("mySortRange").Delete
```

Use the  **Name** property to return or set the text of the name itself. The following example changes the name of the first **Name** object in the active workbook.




```
Names(1).Name = "stock_values"
```

The following example sets the  **Visible** property for the first name assigned to cells A1:B1 on worksheet one.




```
Worksheets(1).Range("a1:b1").Name.Visible = False
```


## Methods



|**Name**|
|:-----|
|[Delete](name-delete-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](name-application-property-excel.md)|
|[Category](name-category-property-excel.md)|
|[CategoryLocal](name-categorylocal-property-excel.md)|
|[Comment](name-comment-property-excel.md)|
|[Creator](name-creator-property-excel.md)|
|[Index](name-index-property-excel.md)|
|[MacroType](name-macrotype-property-excel.md)|
|[Name](name-name-property-excel.md)|
|[NameLocal](name-namelocal-property-excel.md)|
|[Parent](name-parent-property-excel.md)|
|[RefersTo](name-refersto-property-excel.md)|
|[RefersToLocal](name-referstolocal-property-excel.md)|
|[RefersToR1C1](name-referstor1c1-property-excel.md)|
|[RefersToR1C1Local](name-referstor1c1local-property-excel.md)|
|[RefersToRange](name-referstorange-property-excel.md)|
|[ShortcutKey](name-shortcutkey-property-excel.md)|
|[ValidWorkbookParameter](name-validworkbookparameter-property-excel.md)|
|[Value](name-value-property-excel.md)|
|[Visible](name-visible-property-excel.md)|
|[WorkbookParameter](name-workbookparameter-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
