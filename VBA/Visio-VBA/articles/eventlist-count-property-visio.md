---
title: EventList.Count Property (Visio)
keywords: vis_sdr.chm12713330
f1_keywords:
- vis_sdr.chm12713330
ms.prod: visio
api_name:
- Visio.EventList.Count
ms.assetid: c35bd4d3-7b80-71aa-45a7-91e78a51e6eb
ms.date: 06/08/2017
---


# EventList.Count Property (Visio)

Returns the number of objects in a collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents an **EventList** object.


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


