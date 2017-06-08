---
title: Colors.Count Property (Visio)
keywords: vis_sdr.chm12313330
f1_keywords:
- vis_sdr.chm12313330
ms.prod: visio
api_name:
- Visio.Colors.Count
ms.assetid: ade1336b-fc3f-ef2d-0365-914a480260ad
ms.date: 06/08/2017
---


# Colors.Count Property (Visio)

Returns the number of objects in a collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents a **Colors** object.


### Return Value

Long


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


