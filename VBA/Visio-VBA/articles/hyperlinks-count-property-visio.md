---
title: Hyperlinks.Count Property (Visio)
keywords: vis_sdr.chm15613330
f1_keywords:
- vis_sdr.chm15613330
ms.prod: visio
api_name:
- Visio.Hyperlinks.Count
ms.assetid: c460e4ee-679c-b6ed-3c60-e0735bc81815
ms.date: 06/08/2017
---


# Hyperlinks.Count Property (Visio)

Returns the number of objects in a collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents a **Hyperlinks** object.


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


