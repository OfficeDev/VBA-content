---
title: Window.Document Property (Visio)
keywords: vis_sdr.chm11613430
f1_keywords:
- vis_sdr.chm11613430
ms.prod: visio
api_name:
- Visio.Window.Document
ms.assetid: 305471a6-6497-34b4-dfd5-ff37ccb59fff
ms.date: 06/08/2017
---


# Window.Document Property (Visio)

Gets the  **Document** object that is associated with an object. Read-only.


## Syntax

 _expression_ . **Document**

 _expression_ A variable that represents a **Window** object.


### Return Value

Document


## Remarks

The  **Document** property of a docked stencil window returns a **Document** object for the stencil that is currently at the top of the window. If another stencil replaces the first in the top position, the first stencil's document is closed so the reference to it becomes invalid. For best results, assume that document references to docked stencils are not persistent.

If a  **Window** object shows no documents are open, no document is returned and no exception is raised. Your solution should check for **Nothing** returned after retrieving the **Document** property of a **Window** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Document** property of various objects to retrieve data about those objects, and does the following:




- It adds a  **Document** object to the **Documents** collection and sets several of the **Document** object's properties.
    
- It gets the active window and active page, draws a rectangle on the page, and drops a master on the  **Document** object to provide various objects to work on.
    
- It uses the  **Document** property to get the **Document** object associated with each of these other objects.
    





```vb
 
Public Sub Document_Example() 
 
 Dim vsoDocument As Visio.Document 
 Dim vsoTempDocument As Visio.Document 
 Dim vsoPage As Visio.Page 
 Dim vsoShape As Visio.Shape 
 Dim vsoWindow As Visio.Window 
 Dim vsoMaster As Visio.Master 
 
 'Add a document to the Documents collection. 
 Set vsoDocument = Documents.Add("") 
 
 'Set the title of the document. 
 vsoDocument.Title = "My Document" 
 
 'Get the active window and active page. 
 Set vsoWindow = ActiveWindow 
 Set vsoPage = ActivePage 
 
 'Draw a rectangle on the page. 
 Set vsoShape = vsoPage.DrawRectangle(2, 2, 5, 5) 
 
 'Add a master. 
 Set vsoMaster = vsoDocument.Masters.Add 
 
 'Get the Document object associated with various other objects.'Get the Document object associated with the Window object. 
 Set vsoTempDocument = vsoWindow.Document 
 
 'Get the Title property of the Document object'to verify that this is the same document we added earlier. 
 Debug.Print vsoTempDocument.Title 
 
 'Get the Document object associated with the Page object. 
 Set vsoTempDocument = vsoPage.Document 
 Debug.Print vsoTempDocument.Title 
 
 'Get the Document object associated with the Shape object. 
 Set vsoTempDocument = vsoShape.Document 
 Debug.Print vsoTempDocument.Title 
 
 'Get the Document object associated with the Master object. 
 Set vsoTempDocument = vsoMaster.Document 
 Debug.Print vsoTempDocument.Title 
 
End Sub
```


