---
title: Layers.Count Property (Visio)
keywords: vis_sdr.chm11913330
f1_keywords:
- vis_sdr.chm11913330
ms.prod: visio
api_name:
- Visio.Layers.Count
ms.assetid: 455b9f6f-8dd1-28c6-1272-1f6b73f835db
ms.date: 06/08/2017
---


# Layers.Count Property (Visio)

Returns the number of objects in a collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents a **Layers** object.


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


