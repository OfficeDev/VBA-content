---
title: Application.ActiveDocument Property (Visio)
keywords: vis_sdr.chm10013025
f1_keywords:
- vis_sdr.chm10013025
ms.prod: visio
api_name:
- Visio.Application.ActiveDocument
ms.assetid: 0dbc95b6-4920-4291-55c0-ec0128e7f006
ms.date: 06/08/2017
---


# Application.ActiveDocument Property (Visio)

Returns the active  **Document** object, which is the document shown in the active window. Read-only.


## Syntax

 _expression_ . **ActiveDocument**

 _expression_ A variable that represents an **Application** object.


### Return Value

Document


## Remarks

When no documents are open, there is no active document and the  **ActiveDocument** property returns the value **Nothing** and does not raise an exception.

If your code is in the Microsoft Visual Basic project of a Visio document, the  **ActiveDocument** property often, but not necessarily, returns a reference to the **ThisDocument** object, a class module in the Visual Basic project of every Microsoft Visio document. If the **ThisDocument** object is shown in the active window, the **ActiveDocument** object and the **ThisDocument** object refer to the same document. When the **ThisDocument** object is referenced from code in a project, it returns a reference to the project's **Document** object.

Whether you use the  **ActiveDocument** object or the **ThisDocument** object depends on the purpose of your code.

You can compare the result returned by the  **ActiveDocument** property with the value **Nothing** to determine if a document is active. If the value of the **Application.Documents.Count** property is greater than zero, at least one document is open and active.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVApplication.ActiveDocument**
    

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows two safe ways to get an active document (if one exists). In each case, it prints the name of the active document in the  **Immediate** window. The code gets the active document without qualification from the Visio global object, which is automatically available to VBA code that is part of the VBA project of a Visio document.


```vb
 
Public Sub ActiveDocument_Example() 
    
    Dim vsoDocument As Document  
 
    'First method 
    If Documents.Count > 0 Then 
        Set vsoDocument = ActiveDocument  
        Debug.Print vsoDocument.Name  
    Else 
        Debug.Print "No active document."  
    End If   
 
    'Second method 
    If Not(ActiveDocument Is Nothing)  Then 
        Set vsoDocument = ActiveDocument  
        Debug.Print vsoDocument.Name  
    Else 
        Debug.Print "No active document."  
    End If 
   
End Sub
```


