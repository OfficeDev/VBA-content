---
title: Document.Open Event (Publisher)
keywords: vbapb10.chm285212673
f1_keywords:
- vbapb10.chm285212673
ms.prod: publisher
api_name:
- Publisher.Document.Open
ms.assetid: 43108d1d-d101-8a07-943e-c9b8dbadcbfd
ms.date: 06/08/2017
---


# Document.Open Event (Publisher)

Occurs when a publication is opening.


## Syntax

 _expression_. **Open**

 _expression_A variable that represents a  **Document** object.


## Remarks

To access the  **Document** object events, declare a **Document** object variable in the General Declarations section of a class module, then set the variable equal to the **Document** object for which you want to access events.

For more information about using events with the  **Document** object, see [Using Events with the Document Object](using-events-with-the-document-object-publisher.md).


## Example

This example displays a message when a publication is opened. (The procedure can be stored in the  **ThisDocument** module of a publication.)


```vb
Private Sub Document_Open() 
 MsgBox "This publication is copyrighted." 
End Sub
```


