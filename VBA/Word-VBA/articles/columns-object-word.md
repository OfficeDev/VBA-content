---
title: Columns Object (Word)
ms.prod: word
ms.assetid: 7c2d1353-cbc4-a162-83a1-6cac1300266f
ms.date: 06/08/2017
---


# Columns Object (Word)

A collection of  **[Column](column-object-word.md)** objects that represent the columns in a table.


## Remarks

Use the  **Columns** property of a **[Range](range-object-word.md)**, **[Selection](selection-object-word.md)**, or **[Table](table-object-word.md)** object to return a **Columns** collection. The following example displays the number of **Column** objects in the **Columns** collection for the first table in the active document.


```
MsgBox ActiveDocument.Tables(1).Columns.Count
```

The following example creates a table with six columns and three rows and then formats each column with a progressively larger (darker) shading percentage.




```
Set myTable = ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=6) 
For Each col In myTable.Columns 
 col.Shading.Texture = 2 + i 
 i = i + 1 
Next col
```

Use the  **[Add](columns-add-method-word.md)** method to add a column to a table. The following example adds a column to the first table in the active document, and then it makes the column widths equal.




```
If ActiveDocument.Tables.Count >= 1 Then 
 Set myTable = ActiveDocument.Tables(1) 
 myTable.Columns.Add BeforeColumn:=myTable.Columns(1) 
 myTable.Columns.DistributeWidth 
End If
```

Use  **Columns** (Index), where Index is the index number, to return a single **Column** object. The index number represents the position of the column in the **Columns** collection (counting from left to right). The following example selects the first column in the first table.




```
ActiveDocument.Tables(1).Columns(1).Select
```


## Methods



|**Name**|
|:-----|
|[Add](columns-add-method-word.md)|
|[AutoFit](columns-autofit-method-word.md)|
|[Delete](columns-delete-method-word.md)|
|[DistributeWidth](columns-distributewidth-method-word.md)|
|[Item](columns-item-method-word.md)|
|[Select](columns-select-method-word.md)|
|[SetWidth](columns-setwidth-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](columns-application-property-word.md)|
|[Borders](columns-borders-property-word.md)|
|[Count](columns-count-property-word.md)|
|[Creator](columns-creator-property-word.md)|
|[First](columns-first-property-word.md)|
|[Last](columns-last-property-word.md)|
|[NestingLevel](columns-nestinglevel-property-word.md)|
|[Parent](columns-parent-property-word.md)|
|[PreferredWidth](columns-preferredwidth-property-word.md)|
|[PreferredWidthType](columns-preferredwidthtype-property-word.md)|
|[Shading](columns-shading-property-word.md)|
|[Width](columns-width-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
