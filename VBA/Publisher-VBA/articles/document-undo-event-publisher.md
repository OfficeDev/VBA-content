---
title: Document.Undo Event (Publisher)
keywords: vbapb10.chm285212678
f1_keywords:
- vbapb10.chm285212678
ms.prod: publisher
api_name:
- Publisher.Document.Undo
ms.assetid: 9789e469-dc84-a0b7-ffe0-405d4e7ad861
ms.date: 06/08/2017
---


# Document.Undo Event (Publisher)

Occurs when a user undoes the last action performed.


## Syntax

 _expression_. **Undo**

 _expression_A variable that represents a  **Document** object.


## Remarks

The  **Undo** event occurs immediately after the action is undone.

If multiple actions are undone, the  **Undo** event occurs only once, after all the actions are undone.

For more information about using events with the  **Document** object, see [Using Events with the Document Object](using-events-with-the-document-object-publisher.md).


## Example

This example displays a message when the user clicks  **Undo** on the **Standard** toolbar or selects **Undo** from the **Edit** menu. For this routine to work with the current publication, you must put it in the ThisDocument module.


```vb
Private Sub DocPub_Undo() 
 MsgBox "Your last action has been reversed." 
End Sub
```

To trap this event from a non-Microsoft Publisher project, you must place the following code in the General Declarations section of your module and run the InitiatePubApp routine.




```vb
Private WithEvents DocPub As Publisher.Document 
 
Sub InitiatePubApp() 
 Set DocPub = Publisher.ActiveDocument 
End Sub
```


