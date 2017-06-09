---
title: Rows.Add Method (Word)
keywords: vbawd10.chm155975780
f1_keywords:
- vbawd10.chm155975780
ms.prod: word
api_name:
- Word.Rows.Add
ms.assetid: d84286cb-42b5-a717-f152-0d9c3f1c6d9c
ms.date: 06/08/2017
---


# Rows.Add Method (Word)

Returns a  **Row** object that represents a row added to a table.


## Syntax

 _expression_ . **Add**( **_BeforeRow_** )

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BeforeRow_|Optional| **Variant**|A  **Row** object that represents the row that will appear immediately below the new row.|

### Return Value

Row


## Example

This example inserts a new row before the first row in the selection.


```vb
Sub AddARow() 
 If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows.Add BeforeRow:=Selection.Rows(1) 
 End If 
End Sub
```

This example adds a row to the first table and then inserts the text Cell into this row.




```vb
Sub CountCells() 
 Dim tblNew As Table 
 Dim rowNew As Row 
 Dim celTable As Cell 
 Dim intCount As Integer 
 
 intCount = 1 
 Set tblNew = ActiveDocument.Tables(1) 
 Set rowNew = tblNew.Rows.Add(BeforeRow:=tblNew.Rows(1)) 
 For Each celTable In rowNew.Cells 
 celTable.Range.InsertAfter Text:="Cell " &; intCount 
 intCount = intCount + 1 
 Next celTable 
End Sub
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

