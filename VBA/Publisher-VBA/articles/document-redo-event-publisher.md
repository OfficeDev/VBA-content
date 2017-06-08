---
title: Document.Redo Event (Publisher)
keywords: vbapb10.chm285212679
f1_keywords:
- vbapb10.chm285212679
ms.prod: publisher
api_name:
- Publisher.Document.Redo
ms.assetid: c00db13d-1c03-2536-8923-bd7d9393fee2
ms.date: 06/08/2017
---


# Document.Redo Event (Publisher)

Occurs when reversing the last action that was undone.


## Syntax

 _expression_. **Redo**

 _expression_A variable that represents a  **Document** object.


## Remarks

The  **Redo** event occurs immediately after the action is redone.

If multiple actions are redone, the  **Redo** event occurs only once, after all the actions are complete.

For more information about using events with the  **Document** object, see [Using Events with the Document Object](using-events-with-the-document-object-publisher.md).


## Example

This example displays a message when a user clicks  **Undo** on the **Standard** toolbar or selects **Redo** from the **Edit** menu. For this routine to work with the current publication, you must put it in the ThisDocument module.


```vb
Private Sub DocPub_Redo() 
 MsgBox "Your last undo has been reversed." 
End Sub
```

To trap this event from a non-Microsoft Publisher project, you must place the following code in the General Declarations section of your module and run the InitiatePubApp routine.




```vb
Private WithEvents DocPub As Publisher.Document 
 
Sub InitiatePubApp() 
 Set DocPub = Publisher.ActiveDocument 
End Sub
```


