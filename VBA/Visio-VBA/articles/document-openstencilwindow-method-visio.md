---
title: Document.OpenStencilWindow Method (Visio)
keywords: vis_sdr.chm10516420
f1_keywords:
- vis_sdr.chm10516420
ms.prod: visio
api_name:
- Visio.Document.OpenStencilWindow
ms.assetid: 70c3720b-b88d-4859-684b-5c7ae9c868ea
ms.date: 06/08/2017
---


# Document.OpenStencilWindow Method (Visio)

Opens a stencil window that shows the masters in the document.


## Syntax

 _expression_ . **OpenStencilWindow**

 _expression_ A variable that represents a **Document** object.


### Return Value

Window


## Remarks

If the document's stencil is already displayed in a stencil window, the  **OpenStencilWindow** method activates that window rather than opening another window.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **OpenStencilWindow** method to open the Document Stencil window.


```vb
 
Public Sub OpenStencilWindow_Example() 
 
 Dim vsoStencilWindow as Visio.Window 
 
 'Open the Document Stencil window. 
 Set vsoStencilWindow = ThisDocument.OpenStencilWindow 
 
End Sub
```


