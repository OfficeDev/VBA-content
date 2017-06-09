---
title: Working with Tables
ms.prod: word
ms.assetid: cf0858b7-6b39-4c90-552e-edb695b5cda3
ms.date: 06/08/2017
---


# Working with Tables

This topic includes Visual Basic examples related to the following tasks:


-  [Creating a table, inserting text, and applying formatting](#Creating)
    
-  [Inserting text into a table cell](#Inserting)
    
-  [Returning text from a table cell without returning the end of cell marker](#Returning1)
    
-  [Converting text to a table](#Converting)
    
-  [Returning the contents of each table cell](#Returning2)
    
-  [Copying all tables in the active document into a new document](#Copying)
    

## Creating a table, inserting text, and applying formatting

The following example inserts a four-column, three-row table at the beginning of the active document. The  **For Each...Next** structure is used to step through each cell in the table. Within the **For Each...Next** structure, the **[InsertAfter](range-insertafter-method-word.md)** method of the **[Range](range-object-word.md)** object is used to add text to the table cells (Cell 1, Cell 2, and so on).


```vb
Sub CreateNewTable() 
 Dim docActive As Document 
 Dim tblNew As Table 
 Dim celTable As Cell 
 Dim intCount As Integer 
 
 Set docActive = ActiveDocument 
 Set tblNew = docActive.Tables.Add( _ 
 Range:=docActive.Range(Start:=0, End:=0), NumRows:=3, _ 
 NumColumns:=4) 
 intCount = 1 
 
 For Each celTable In tblNew.Range.Cells 
 celTable.Range.InsertAfter "Cell " &; intCount 
 intCount = intCount + 1 
 Next celTable 
 
 tblNew.AutoFormat Format:=wdTableFormatColorful2, _ 
 ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True 
End Sub
```


## Inserting text into a table cell

The following example inserts text into the first cell of the first table in the active document. The  **[Cell](table-cell-method-word.md)** method returns a single  **Cell** object. The **Range**property returns a  **Range** object. The **[Delete](range-delete-method-word.md)** method is used to delete the existing text and the  **InsertAfter**method inserts the "Cell 1,1" text.


```vb
Sub InsertTextInCell() 
 If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Cell(Row:=1, Column:=1).Range 
 .Delete 
 .InsertAfter Text:="Cell 1,1" 
 End With 
 End If 
End Sub
```


## Returning text from a table cell without returning the end of cell marker

The following example returns and displays the contents of each cell in the first row of the first document table.


```vb
Sub ReturnTableText() 
 Dim tblOne As Table 
 Dim celTable As Cell 
 Dim rngTable As Range 
 
 Set tblOne = ActiveDocument.Tables(1) 
 For Each celTable In tblOne.Rows(1).Cells 
 Set rngTable = ActiveDocument.Range(Start:=celTable.Range.Start, _ 
 End:=celTable.Range.End - 1) 
 MsgBox rngTable.Text 
 Next celTable 
End Sub
```


```vb
Sub ReturnCellText() 
 Dim tblOne As Table 
 Dim celTable As Cell 
 Dim rngTable As Range 
 
 Set tblOne = ActiveDocument.Tables(1) 
 For Each celTable In tblOne.Rows(1).Cells 
 Set rngTable = celTable.Range 
 rngTable.MoveEnd Unit:=wdCharacter, Count:=-1 
 MsgBox rngTable.Text 
 Next celTable 
End Sub
```


## Converting existing text to a table

The following example inserts tab-delimited text at the beginning of the active document and then converts the text to a table.


```vb
Sub ConvertExistingText() 
 With Documents.Add.Content 
 .InsertBefore "one" &; vbTab &; "two" &; vbTab &; "three" &; vbCr 
 .ConvertToTable Separator:=Chr(9), NumRows:=1, NumColumns:=3 
 End With 
End Sub
```


## Returning the contents of each table cell

The following example defines an array equal to the number of cells in the first document table (assuming  **Option Base 1**). The  **For Each...Next** structure is used to return the contents of each table cell and assign the text to the corresponding array element.


```vb
Sub ReturnCellContentsToArray() 
 Dim intCells As Integer 
 Dim celTable As Cell 
 Dim strCells() As String 
 Dim intCount As Integer 
 Dim rngText As Range 
 
 If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Range 
 intCells = .Cells.Count 
 ReDim strCells(intCells) 
 intCount = 1 
 For Each celTable In .Cells 
 Set rngText = celTable.Range 
 rngText.MoveEnd Unit:=wdCharacter, Count:=-1 
 strCells(intCount) = rngText 
 intCount = intCount + 1 
 Next celTable 
 End With 
 End If 
End Sub
```


## Copying all tables in the active document into a new document

This example copies the tables from the current document into a new document.


```vb
Sub CopyTablesToNewDoc() 
 Dim docOld As Document 
 Dim rngDoc As Range 
 Dim tblDoc As Table 
 
 If ActiveDocument.Tables.Count >= 1 Then 
 Set docOld = ActiveDocument 
 Set rngDoc = Documents.Add.Range(Start:=0, End:=0) 
 For Each tblDoc In docOld.Tables 
 tblDoc.Range.Copy 
 With rngDoc 
 .Paste 
 .Collapse Direction:=wdCollapseEnd 
 .InsertParagraphAfter 
 .Collapse Direction:=wdCollapseEnd 
 End With 
 Next 
 End If 
End Sub
```


