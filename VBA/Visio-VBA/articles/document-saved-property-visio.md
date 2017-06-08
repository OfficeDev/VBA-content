---
title: Document.Saved Property (Visio)
keywords: vis_sdr.chm10514285
f1_keywords:
- vis_sdr.chm10514285
ms.prod: visio
api_name:
- Visio.Document.Saved
ms.assetid: de3141f6-eda9-a62b-847c-e946966fae6b
ms.date: 06/08/2017
---


# Document.Saved Property (Visio)

Indicates whether a document has any unsaved changes. Read/write.


## Syntax

 _expression_ . **Saved**

 _expression_ A variable that represents a **Document** object.


### Return Value

Boolean


## Remarks

Use caution when setting the  **Saved** property for a document to **True** . If you set the **Saved** property to **True** and a user, or another program, makes changes to the document before it is closed, those changes will be lost—Microsoft Visio does not provide a prompt to save the document.

A document that contains embedded or linked OLE objects may report itself as unsaved even if the document's  **Saved** property is set to **True** .


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Saved** property to determine whether a document has any unsaved changes. It also shows how to set the **Saved** property. Before running this macro, change _path_ to the location where you want to save the drawing, and change _filename_ to the name you'd like to assign the file.


```vb
 
Public Sub Saved_Example() 
 
 Dim vsoDocument1 As Visio.Document 
 Dim vsoDocument2 As Visio.Document 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 
 Set vsoPage = ThisDocument.Pages(1) 
 Set vsoShape = vsoPage.DrawOval(2.5, 7, 3.5, 9) 
 
 'Use the SaveAs method to save the document for the first time. 
 ThisDocument.SaveAs "path\filename .vsd" 
 
 'Use the Saved property to verify that the document was saved. 
 'Saved returns True (-1). 
 Debug.Print ThisDocument.Saved 
 
 'Force a change to the document by adding a shape. 
 Set vsoShape = vsoPage.DrawOval(4, 7, 5, 9) 
 
 'Use the Saved property to verify that the document changed 
 'since the last time is was saved. 
 'Saved returns False (0) 
 Debug.Print ThisDocument.Saved 
 
 'Use the Save method to save any new changes. 
 ThisDocument.Save 
 
 'Use the Saved property again to verify that 
 'the document was saved. Saved returns True (-1). 
 Debug.Print ThisDocument.Saved 
 
 'The Saved property can also be set. For example, change 
 'the document again so that the Saved property becomes False. 
 Set vsoShape = vsoPage.DrawRectangle(1, 1, 7, 7) 
 
 'Set the Saved property to True. 
 'Setting the Saved property to True does not save the document. 
 ThisDocument.Saved = True 
 
 'Close the document and then reopen it. Note that 
 'the rectangle was not saved. 
 Set vsoDocument1 = ThisDocument 
 vsoDocument1.Close 
 Set vsoDocument1 = Documents.Open("path\filename .vsd") 
 
End Sub
```


