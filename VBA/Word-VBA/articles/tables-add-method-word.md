---
title: Tables.Add Method (Word)
keywords: vbawd10.chm156041416
f1_keywords:
- vbawd10.chm156041416
ms.prod: word
api_name:
- Word.Tables.Add
ms.assetid: 127b5f74-876f-1307-5d25-a04c99debd6b
ms.date: 06/08/2017
---


# Tables.Add Method (Word)

Returns a  **Table** object that represents a new, blank table added to a document.


## Syntax

 _expression_ . **Add**( **_Range_** , **_NumRows_** , **_NumColumns_** , **_DefaultTableBehavior_** , **_AutoFitBehavior_** )

 _expression_ Required. A variable that represents a **[Tables](tables-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range where you want the table to appear. The table replaces the range, if the range isn't collapsed.|
| _NumRows_|Required| **Long**|The number of rows you want to include in the table.|
| _NumColumns_|Required| **Long**|The number of columns you want to include in the table.|
| _DefaultTableBehavior_|Optional| **Variant**|Sets a value that specifies whether Microsoft Word automatically resizes cells in tables to fit the cells? contents (AutoFit). Can be either of the following constants:  **wdWord8TableBehavior** (AutoFit disabled) or **wdWord9TableBehavior** (AutoFit enabled). The default constant is **wdWord8TableBehavior** .|
| _AutoFitBehavior_|Optional| **Variant**|Sets the AutoFit rules for how Word sizes tables. Can be one of the  **WdAutoFitBehavior** constants.|

### Return Value

Table


## Example

This example adds a blank table with three rows and four columns at the beginning of the active document.


```vb
Set myRange = ActiveDocument.Range(0, 0) 
ActiveDocument.Tables.Add Range:=myRange, NumRows:=3, NumColumns:=4
```

This example adds a new, blank table with six rows and ten columns at the end of the active document




```vb
Set MyRange = ActiveDocument.Content 
MyRange.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Tables.Add Range:=MyRange, NumRows:=6, _ 
 NumColumns:=10
```

This example adds a table with three rows and five columns to a new document and then inserts data into each cell in the table.




```vb
Sub NewTable() 
 Dim docNew As Document 
 Dim tblNew As Table 
 Dim intX As Integer 
 Dim intY As Integer 
 
 Set docNew = Documents.Add 
 Set tblNew = docNew.Tables.Add(Selection.Range, 3, 5) 
 With tblNew 
 For intX = 1 To 3 
 For intY = 1 To 5 
 .Cell(intX, intY).Range.InsertAfter "Cell: R" &; intX &; ", C" &; intY 
 Next intY 
 Next intX 
 .Columns.AutoFit 
 End With 
End Sub
```


## See also


#### Concepts


[Tables Collection Object](tables-object-word.md)

