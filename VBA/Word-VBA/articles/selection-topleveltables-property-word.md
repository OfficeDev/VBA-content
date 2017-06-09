---
title: Selection.TopLevelTables Property (Word)
keywords: vbawd10.chm158663662
f1_keywords:
- vbawd10.chm158663662
ms.prod: word
api_name:
- Word.Selection.TopLevelTables
ms.assetid: 7ab1b2a3-85a8-8892-53b9-dc85ff747078
ms.date: 06/08/2017
---


# Selection.TopLevelTables Property (Word)

Returns a  **[Tables](tables-object-word.md)** collection that represents the tables at the outermost nesting level in the current selection. Read-only.


## Syntax

 _expression_ . **TopLevelTables**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

This method returns a collection containing only those tables at the outermost nesting level within the context of the current selection. These tables may not be at the outermost nesting level within the entire set of nested tables.

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example creates a new document, creates a nested table with three levels, and then fills the first cell of each table with its nesting level. The example selects the second column of the second-level table and then selects the first of the top-level tables in this selection. The innermost table is selected, even though it isn't a top-level table within the context of the entire set of nested tables.


```vb
Documents.Add 
ActiveDocument.Tables.Add Selection.Range, _ 
 3, 3, wdWord9TableBehavior, wdAutoFitContent 
With ActiveDocument.Tables(1).Range 
 .Copy 
 .Cells(1).Range.Text = .Cells(1).NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = .Cells(1).NestingLevel 
 .Cells(5).Range.PasteAsNestedTable 
 With .Cells(5).Tables(1).Range 
 .Cells(1).Range.Text = _ 
 .Cells(1).NestingLevel 
 End With 
 .Columns(2).Select 
 Selection.TopLevelTables(1).Select 
 End With 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

