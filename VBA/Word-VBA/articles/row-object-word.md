---
title: Row Object (Word)
keywords: vbawd10.chm2384
f1_keywords:
- vbawd10.chm2384
ms.prod: word
api_name:
- Word.Row
ms.assetid: 38a05858-829a-ea5c-ce63-7f7343bf7b88
ms.date: 06/08/2017
---


# Row Object (Word)

Represents a row in a table. The  **Row** object is a member of the **[Rows](rows-object-word.md)** collection. The **Rows** collection includes all the rows in the specified selection, range, or table.


## Remarks

Use  **Rows** (Index), where Index is the index number, to return a single **Row** object. The index number represents the position of the row in the selection, range, or table. The following example deletes the first row in the first table in the active document.


```vb
ActiveDocument.Tables(1).Rows(1).Delete
```

Use the  **Add** method to add a row to a table. The following example inserts a row before the first row in the selection.




```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows.Add BeforeRow:=Selection.Rows(1) 
End If
```

Use the  **Cells** property to modify the individual cells in a **Row** object. The following example adds a table to the selection and then inserts numbers into each cell in the second row of the table.




```vb
Selection.Collapse Direction:=wdCollapseEnd 
If Selection.Information(wdWithInTable) = False Then 
 Set myTable = _ 
 ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=3, NumColumns:=5) 
 For Each aCell In myTable.Rows(2).Cells 
 i = i + 1 
 aCell.Range.Text = i 
 Next aCell 
End If
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


