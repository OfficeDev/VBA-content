---
title: UndoRecord.EndCustomRecord Method (Word)
keywords: vbawd10.chm56098818
f1_keywords:
- vbawd10.chm56098818
ms.prod: word
api_name:
- Word.UndoRecord.EndCustomRecord
ms.assetid: af11d231-f799-d592-2bc5-de08030b41e4
ms.date: 06/08/2017
---


# UndoRecord.EndCustomRecord Method (Word)

Completes the creation of a custom undo record.


## Syntax

 _expression_ . **EndCustomRecord**

 _expression_ A variable that represents an **[UndoRecord](undorecord-object-word.md)** object.


## Remarks

You use the [UndoRecord.StartCustomRecord](undorecord-startcustomrecord-method-word.md) to initiate the creation of a custom undo record. To complete the creation of a custom undo record, you use the **EndCustomRecord** method.


## Example

The following code example creates a custom undo record.


```vb
Sub TestUndo() 
Dim objUndo As UndoRecord 
 
Set objUndo = Application.UndoRecord 
objUndo.StartCustomRecord ("My Custom Undo") 
    'Add some actions here 
objUndo.EndCustomRecord 
     
End Sub
```


## See also


#### Concepts


[UndoRecord Object](undorecord-object-word.md)
#### Other resources


[Working with the UndoRecord Object](http://msdn.microsoft.com/library/e9df1047-5a1a-91da-3673-7e64b668552d%28Office.15%29.aspx)


