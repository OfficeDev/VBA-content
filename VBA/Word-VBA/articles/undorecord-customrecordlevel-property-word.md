---
title: UndoRecord.CustomRecordLevel Property (Word)
keywords: vbawd10.chm56098821
f1_keywords:
- vbawd10.chm56098821
ms.prod: word
api_name:
- Word.UndoRecord.CustomRecordLevel
ms.assetid: e0636c02-b1fb-2f88-c8a5-b52c88b65530
ms.date: 06/08/2017
---


# UndoRecord.CustomRecordLevel Property (Word)

Returns a  **Long** that specifies the number of custom undo action calls that are currently active. Read-only.


## Syntax

 _expression_ . **CustomRecordLevel**

 _expression_ A variable that represents a **[UndoRecord](undorecord-object-word.md)** object.


## Remarks

If no custom undo action is active, this property is set to 0.


## Example

The following code example verifies that a custom undo record is currently recording. If not, the code creates a custom undo record. Finally, the code verifies that any custom undo action calls are active. If so, a message is printed to the Debug window.


```vb
Dim objUndo As UndoRecord 
 
Sub MyFunction() 
 Set objUndo = Application.UndoRecord 
 
 ' Verify that a custom undo record is already being recorded, and if not, start one 
 If objUndo.IsRecordingCustomRecord = False Then 
 objUndo.StartCustomRecord("New Undo Record") 
 End If 
 ' Add some actions here. 
 objUndo.EndCustomRecord 
 
 ' Verify that any custom undo action calls are currently active. 
 If objUndo.CustomRecordLevel > 0 Then 
 Debug.Print "An undo record call was not closed!" 
 End If 
End Sub
```


## See also


#### Concepts


[UndoRecord Object](undorecord-object-word.md)

