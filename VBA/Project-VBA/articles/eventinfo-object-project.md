---
title: EventInfo Object (Project)
keywords: vbapj.chm131286
f1_keywords:
- vbapj.chm131286
ms.prod: project-server
api_name:
- Project.EventInfo
ms.assetid: 97a51ee0-f7eb-5215-0686-1944c537e8fc
ms.date: 06/08/2017
---


# EventInfo Object (Project)

Represents cancellation information for an event.
 


## Remarks

The  **EventInfo** object has one **Boolean** property, named **Cancel**. Project uses the **EventInfo** object instead of the _Cancel_ parameter that is used for events in some previous versions of Project.
 

 

## Example

The following event handler examines new resource assignments and cancels them if they are for the specified resource.
 

 

```
Private Sub App_ProjectBeforeAssignmentChange2(ByVal asg As Assignment, ByVal Field As PjAssignmentField, _ 
    ByVal NewVal As Variant, EventInfo As Object) 
 
    If Field = pjAssignmentResourceName And NewVal = "Lisa Jones" Then 
        MsgBox "Lisa is no longer available for assignment!" 
        EventInfo.Cancel = True 
    End If 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Cancel](eventinfo-cancel-property-project.md)|

