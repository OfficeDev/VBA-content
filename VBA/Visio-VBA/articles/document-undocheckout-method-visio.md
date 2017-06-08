---
title: Document.UndoCheckOut Method (Visio)
keywords: vis_sdr.chm10560090
f1_keywords:
- vis_sdr.chm10560090
ms.prod: visio
api_name:
- Visio.Document.UndoCheckOut
ms.assetid: 7b6a67ae-2acd-217f-42e0-f8aced97ac96
ms.date: 06/08/2017
---


# Document.UndoCheckOut Method (Visio)

Closes a Microsoft Visio document checked out from a Microsoft SharePoint Server site, deletes the local copy of the document, discarding any changes, undoes the checkout, and then reopens the document.


## Syntax

 _expression_ . **UndoCheckOut**

 _expression_ An expression that returns a **Document** object.


### Return Value

Nothing


## Remarks

Calling the  **UndoCheckOut** method is the equivalent of clicking **Discard Check Out** on the **Check In** drop-down menu (click the **File** tab, and then click **Info**).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **UndoCheckOut** method to undo the checkout of the active document from a SharePoint server. Before running this macro, check out a Visio document from a SharePoint Server site.


```vb
Public Sub UndoCheckOut_Example 
    Visio.ActiveDocument.UndoCheckOut 
End Sub
```


