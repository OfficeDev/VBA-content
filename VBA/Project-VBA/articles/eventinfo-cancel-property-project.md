---
title: EventInfo.Cancel Property (Project)
keywords: vbapj.chm131085
f1_keywords:
- vbapj.chm131085
ms.prod: project-server
api_name:
- Project.EventInfo.Cancel
ms.assetid: 2bd3a795-9a8f-8cdb-5358-a22487610a72
ms.date: 06/08/2017
---


# EventInfo.Cancel Property (Project)

In an event handler, the  **Cancel** property gets or sets a value that specifies whether the operation that triggered the event should continue. If **True**, the operation is canceled. Read/write **Boolean**.


## Syntax

 _expression_. **Cancel**

 _expression_ A variable that represents an **EventInfo** object.


## Remarks

The default value of the  **Cancel** property is **False** when an event occurs. Set **Cancel** to **True** to cancel an operation.


## Example

The following event handler examines new resource assignments and cancels them if they are for the specified resource.


```vb
Private Sub App_ProjectBeforeAssignmentChange2(ByVal asg As Assignment, ByVal Field As PjAssignmentField, _ 
 ByVal NewVal As Variant, EventInfo As Object) 
 
 If Field = pjAssignmentResourceName And NewVal = "Lisa Jones" Then 
 MsgBox "Lisa is no longer available for assignment!" 
 EventInfo.Cancel = True 
 End If 
End Sub
```


