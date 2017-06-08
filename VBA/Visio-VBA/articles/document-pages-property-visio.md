---
title: Document.Pages Property (Visio)
keywords: vis_sdr.chm10513995
f1_keywords:
- vis_sdr.chm10513995
ms.prod: visio
api_name:
- Visio.Document.Pages
ms.assetid: db81b42f-dfd7-c4dc-a520-b1927cd1e737
ms.date: 06/08/2017
---


# Document.Pages Property (Visio)

Returns the  **Pages** collection for a document. Read-only.


## Syntax

 _expression_ . **Pages**

 _expression_ A variable that represents a **Document** object.


### Return Value

Pages


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVDocument.Pages**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Pages** property to print the names of a document's pages.


```vb
 
Public Sub Pages_Example() 
  
    Dim intCounter As Integer 
    Dim vsoDocument As Visio.Document  
    Dim vsoPages As Visio.Pages  
 
    'Get the Pages collection for the active document.  
    Set vsoPages = ActiveDocument.Pages 
  
    Debug.Print "Page names for document: "; ActiveDocument.Name 
  
    'Iterate through the pages and print the page name  
    'in the Immediate window.  
    For intCounter = 1 To vsoPages.Count  
        Debug.Print " "; vsoPages.Item(intCounter).Name  
    Next intCounter  
 
End Sub
```


