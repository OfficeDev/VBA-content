---
title: Document.XMLAfterInsert Event (Word)
keywords: vbawd10.chm4001008
f1_keywords:
- vbawd10.chm4001008
ms.prod: word
api_name:
- Word.Document.XMLAfterInsert
ms.assetid: 6858fd64-e96b-308e-53eb-d79595fabac7
ms.date: 06/08/2017
---


# Document.XMLAfterInsert Event (Word)

Occurs when a user adds a new XML element to a document. If more than one element is added to the document at the same time (for example, when cutting and pasting XML), the event fires for each element that is inserted.


## Syntax

Private Sub  _expression_ _**XMLAfterInsert**( **_NewXMLNode_** , **_InUndoRedo_** )

 _expression_ A variable that represents a **[Document](document-object-word.md)** object that has been declared by using the **WithEvents** keyword in a class module. For information about using events with a **Document** object, see[Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewXMLNode_|Required| **[XMLNode](xmlnode-object-word.md)**|The newly added XML node.|
| _InUndoRedo_|Required| **Boolean**| **True** indicates the action was performed using the **Undo** or **Redo** feature in Microsoft Word.|

## Remarks

If the InUndoRedo parameter is  **True** , never change the XML in a document while the **XMLAfterInsert** and **XMLBeforeDelete** events are running.

If the InUndoRedo parameter is  **False** , you can insert and delete the XML in the document, but be careful that the **XMLAfterInsert** and **XMLBeforeDelete** events will not try to cancel each other out, causing an infinite loop. You can prevent infinite loops by using a global **Boolean** variable and check for that at the beginning of the error handler, as shown in the following example.




```vb
Dim blnIsXMLInsertRunning As Boolean 
 
Private Sub Document_XMLAfterInsert(ByVal DeletedRange As Range, _ 
 ByVal OldXMLNode As XMLNode, ByVal InUndoRedo As Boolean) 
 
 If blnIsXMLInsertRunning = False Then 
 blnIsXMLInsertRunning = True 
 'Insert your event code here. 
 Else 
 Exit Sub 
 End If 
End Sub
```


## Example

The following example validates a newly added node and if the node is not valid, displays a message describing the validation error.


```vb
Private Sub Document_XMLAfterInsert(ByVal NewXMLNode As XMLNode, _ 
 ByVal InUndoRedo As Boolean) 
 
 NewXMLNode.Validate 
 
 If NewXMLNode.ValidationStatus <> wdXMLValidationStatusOK Then 
 MsgBox NewXMLNode.ValidationErrorText 
 End If 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

