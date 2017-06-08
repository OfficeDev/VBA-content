---
title: Using Events with the Document Object
ms.prod: word
ms.assetid: 2b043342-436a-5421-e8af-3c2c49684960
ms.date: 06/08/2017
---


# Using Events with the Document Object

The  **[Document](document-object-word.md)** object supports several events that enable you to respond to the state of a document. You write procedures to respond to these events in the class module named "ThisDocument." Use the following steps to create an event procedure.


1. Under your Normal project or document project in the Project Explorer window, double-click  **ThisDocument**. (In Folder view,  **ThisDocument** is located in the **Microsoft Word Objects** folder.)
    
2. Select  **Document** from the **Object** drop-down list box.An empty subroutine for the **New** event is added to the class module.
    
3. Select an event from the  **Procedure** drop-down list box. An empty subroutine for the selected event is added to the class module.
    
4. Add the Visual Basic instructions you want to run when the event occurs.
    

The following example shows a  **[New](document-new-event-word.md)** event procedure in the Normal project that will run when a new document based on the Normal template is created.




```vb
Private Sub Document_New() 
 MsgBox "New document was created" 
End Sub
```

The following example shows a  **[Close](document-close-event-word.md)** event procedure in a document project that runs only when that document is closed.



```vb
Private Sub Document_Close() 
 MsgBox "Closing the document" 
End Sub
```

Unlike  [auto macros](auto-macros.md), event procedures in the Normal template do not have a global scope. For example, event procedures in the Normal template occur only if the attached template is the Normal template.
If an auto macro exists in a document and the attached template, only the auto macro stored in the document will execute. If an event procedure for a document event exists in a document and its attached template, both event procedures will run.

## Remarks

For information about creating event procedures for the  **[Application](application-object-word.md)** object, see  [Using Events with the Application Object](using-events-with-the-application-object-word.md).


