---
title: Document.WizardAfterChange Event (Publisher)
keywords: vbapb10.chm285212676
f1_keywords:
- vbapb10.chm285212676
ms.prod: publisher
api_name:
- Publisher.Document.WizardAfterChange
ms.assetid: c4ec0950-3a58-1f29-b35f-35db9d87f330
ms.date: 06/08/2017
---


# Document.WizardAfterChange Event (Publisher)

Occurs after the user chooses an option in the wizard pane that changes any of the following settings in the publication: page layout (page size, fold type, orientation, label product), print setup (paper size, print tiling), adding or deleting objects, adding or deleting pages, or object or page formatting (size, position, fill, border, background, default text, text formatting).


## Syntax

 _expression_. **WizardAfterChange**

 _expression_A variable that represents a  **Document** object.


## Remarks

The WizardAfterChange event only occurs once regardless of the scope or number of individual modifications made to the publication.

To access the  **Document** object events, declare a **Document** object variable in the General Declarations section of a class module, then set the variable equal to the **Document** object for which you want to access events.

For more information about using events with the  **Document** object, see [Using Events with the Document Object](using-events-with-the-document-object-publisher.md).


## Example

This example displays a message when a publication is altered using the wizard pane. (The procedure can be stored in the ThisDocument module of a publication.)


```vb
Private Sub Document_WizardAfterChange() 
 MsgBox "Remember to save changes made " _ 
 &; "through the wizard pane!" 
End Sub
```


