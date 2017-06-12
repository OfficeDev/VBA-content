---
title: Pages.Count Property (Visio)
keywords: vis_sdr.chm11013330
f1_keywords:
- vis_sdr.chm11013330
ms.prod: visio
api_name:
- Visio.Pages.Count
ms.assetid: 1e240cc4-07f3-ceb1-7eb3-7a6d5071f630
ms.date: 06/08/2017
---


# Pages.Count Property (Visio)

Returns the number of objects in a collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents a **Pages** object.


### Return Value

Integer


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Count** property to iterate through a **Documents** collection. It displays the names of all the open Microsoft Visio documents in the **Immediate** window.


```vb
 
Public Sub Count_Example() 
 
 Dim intCounter As Integer 
 Dim vsoDocument As Visio.Document 
 
 For intCounter = 1 To Documents.Count 
 'Get the next open document. 
 Set vsoDocument = Documents.Item(intCounter) 
 
 'Print its name in the Immediate window. 
 Debug.Print vsoDocument.Name 
 Next intCounter 
 
End Sub
```


