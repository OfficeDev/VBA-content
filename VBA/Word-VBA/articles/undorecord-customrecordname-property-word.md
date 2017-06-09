---
title: UndoRecord.CustomRecordName Property (Word)
keywords: vbawd10.chm56098820
f1_keywords:
- vbawd10.chm56098820
ms.prod: word
api_name:
- Word.UndoRecord.CustomRecordName
ms.assetid: 97da07e1-3b9f-de7d-c2d8-af6af2bb2374
ms.date: 06/08/2017
---


# UndoRecord.CustomRecordName Property (Word)

Returns a  **String** that specifies the entry that appears on the undo stack when all custom undo actions have completed. Read-only.


## Syntax

 _expression_ . **CustomRecordName**

 _expression_ A variable that represents a **[UndoRecord](undorecord-object-word.md)** object.


## Remarks

If custom undo records are nested within other custom undo records, this property specifies what string appears on the undo stack after all custom undo actions have completed. If multiple calls to the [StartCustomRecord](undorecord-startcustomrecord-method-word.md) method are nested, the string specified by the first call will be returned by this property. If no action is active, the property returns an empty string.


## Example

The following code example creates nested custom undo records. When the code completes, a message about each undo record is inserted into the active document, and "First call" appears as the entry on the undo stack.


 **Note**  To run this code example, place it the code file for  **ThisDocument** in the Visual Basic for Applications Project Explorer.


```vb
Sub WalkUndoRecordStack()
Dim objUndo As UndoRecord
 
'Create UndoRecord object
Set objUndo = Application.UndoRecord
 
'Begin first custom record
objUndo.StartCustomRecord ("First call")
    'Begin nested second custom record
    objUndo.StartCustomRecord ("Second call")
        'Begin nested third undo record
        objUndo.StartCustomRecord ("Third call")
            'Message for the third call is written first to the document
            Me.Content.InsertAfter "Third call. "
            'End third custom record
        objUndo.EndCustomRecord
        'Message for the second call is written second to the document
        Me.Content.InsertAfter "Second call. "
    'End second custom record
    objUndo.EndCustomRecord
    'Message for first call is written third to the document
    Me.Content.InsertAfter "First call. "
'End first custom record
objUndo.EndCustomRecord

Set objUndo = Nothing
End Sub
```


## See also


#### Concepts


[UndoRecord Object](undorecord-object-word.md)

