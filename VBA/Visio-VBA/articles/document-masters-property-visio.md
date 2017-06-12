---
title: Document.Masters Property (Visio)
keywords: vis_sdr.chm10513875
f1_keywords:
- vis_sdr.chm10513875
ms.prod: visio
api_name:
- Visio.Document.Masters
ms.assetid: b139014c-6d7c-ba76-8366-bcacecc5c639
ms.date: 06/08/2017
---


# Document.Masters Property (Visio)

Returns the  **Masters** collection for a document's stencil. Read-only.


## Syntax

 _expression_ . **Masters**

 _expression_ A variable that represents a **Document** object.


### Return Value

Masters


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVDocument.Masters**
    

## Example

This Microsoft Visual Basic for Applications (VBA) program shows how to use the  **Masters** property to print all the master shape names in the current document to the Immediate window.



Before running this macro, open a Microsoft Visio drawing and drag at least one shape from a stencil onto to the drawing page.




```vb
 
Public Sub Masters_Example() 
  
    Dim intCounter As Integer 
    Dim intMasterCount As Integer 
    Dim vsoApplication As Visio.Application  
    Dim vsoCurrentDocument As Visio.Document  
    Dim vsoMasters As Visio.Masters 
 
    Set vsoApplication = GetObject(, "visio.application") 
  
    If vsoApplication Is Nothing Then 
        MsgBox "Microsoft Office Visio is not loaded" 
        Exit Sub   
 
    End If   
 
    Set vsoCurrentDocument = vsoApplication.ActiveDocument  
 
    If vsoCurrentDocument Is Nothing Then 
        MsgBox "No stencil is loaded" 
        Exit Sub   
 
    End If   
 
    Set vsoMasters = vsoCurrentDocument.Masters  
    Debug.Print "Masters in document : "; vsoCurrentDocument.Name  
    intMasterCount = vsoMasters.Count  
 
    If intMasterCount > 0 Then 
        For intCounter = 1 To intMasterCount  
            Debug.Print " "; vsoMasters.Item(intCounter).Name  
        Next intCounter  
    Else 
        Debug.Print " No masters in document"  
    End If   
 
End Sub
```


