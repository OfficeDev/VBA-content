---
title: Document.CanUndoCheckOut Method (Visio)
keywords: vis_sdr.chm10560085
f1_keywords:
- vis_sdr.chm10560085
ms.prod: visio
api_name:
- Visio.Document.CanUndoCheckOut
ms.assetid: aa271635-73ef-b681-364c-49d515fd54cb
ms.date: 06/08/2017
---


# Document.CanUndoCheckOut Method (Visio)

Determines whether a Microsoft Visio document is checked out from a Microsoft SharePoint Server site, so that if it is, the check-out can be subsequently undone.


## Syntax

 _expression_ . **CanUndoCheckOut**

 _expression_ An expression that returns a **Document** object.


### Return Value

Boolean


## Remarks

The  **CanUndoCheckOut** method is similar to the **[Document.CanCheckIn](document-cancheckin-method-visio.md)** method.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **CanUndoCheckOut** method to determine if the checkout of the active document from a SharePoint server site can be undone. Before running this macro, check out a Visio document from a SharePoint Server site.


```vb
Public Sub CanUndoCheckOut_Example 
    
    Dim boolCanUndo As Boolean 
    boolCanUndo = Visio.ActiveDocument.CanUndoCheckOut 
         
    Debug.Print boolCanUndo 
 
End Sub
```


