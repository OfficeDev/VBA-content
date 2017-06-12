---
title: Tables Object (Word)
keywords: vbawd10.chm2381
f1_keywords:
- vbawd10.chm2381
ms.prod: word
ms.assetid: 068a3d0f-0b19-3927-cb0a-7fb0d0fd8e52
ms.date: 06/08/2017
---


# Tables Object (Word)

A collection of  **[Table](table-object-word.md)** objects that represent the tables in a selection, range, or document.


## Remarks

Use the  **Tables** property to return the **Tables** collection. The following example applies a border around each of the tables in the active document.


```vb
For Each aTable In ActiveDocument.Tables 
 aTable.Borders.OutsideLineStyle = wdLineStyleSingle 
 aTable.Borders.OutsideLineWidth = wdLineWidth025pt 
 aTable.Borders.InsideLineStyle = wdLineStyleNone 
Next aTable
```

Use the  **Add** method to add a table at the specified range. The following example adds a 3x4 table at the beginning of the active document.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.Tables.Add Range:=myRange, NumRows:=3, NumColumns:=4
```

Use  **Tables** (Index), where Index is the index number, to return a single **Table** object. The index number represents the position of the table in the selection, range, or document. The following example converts the first table in the active document to text.




```vb
ActiveDocument.Tables(1).ConvertToText Separator:=wdSeparateByTabs
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


