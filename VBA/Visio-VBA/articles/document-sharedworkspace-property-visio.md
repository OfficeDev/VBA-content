---
title: Document.SharedWorkspace Property (Visio)
keywords: vis_sdr.chm10560136
f1_keywords:
- vis_sdr.chm10560136
ms.prod: visio
api_name:
- Visio.Document.SharedWorkspace
ms.assetid: 100d635c-2b2a-4ba3-0490-bc4a4c4efb8c
ms.date: 06/08/2017
---


# Document.SharedWorkspace Property (Visio)

Returns a Microsoft Office  **SharedWorkspace** object that provides access to the Office Document Workspace object model. Read-only.


## Syntax

 _expression_ . **SharedWorkspace**

 _expression_ A variable that represents a **Document** object.


### Return Value

Object


## Remarks

The Office Document Workspace object model provides a way to put documents into a shared workspace and manipulate Microsoft SharePoint data such as people, tasks, links, and related files.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SharedWorkspace** property to get a **SharedWorkspace** object and create a new shared document workspace that has the same name as the default document, at the default location.


```vb
Public Sub SharedWorkspace_Example 
 
 Dim vsoSharedWorkspace As SharedWorkspace 
 Set vsoSharedWorkspace = ActiveDocument.SharedWorkspace 
 vsoSharedWorkspace.CreateNew ("") 
 
End Sub
```


