---
title: Resource.MaxUnits Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.MaxUnits
ms.assetid: 1c698f41-9bd2-8673-af5c-6dce48a75511
ms.date: 06/08/2017
---


# Resource.MaxUnits Property (Project)

Gets or sets the maximum percent availability of the resource. Read/write  **Variant**.


## Syntax

 _expression_. **MaxUnits**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The percent availability is specified in the  **Units** column of the current row of the **Resource Availability** grid in the **Resource Information** dialog box. The current row is that where the date range between the **Available From** and **Available To** columns includes the current date.

The  **MaxUnits** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.


## Example

The following example sets the maximum units of each resource in the active project to a number specified by the user.


```vb
Sub SetDefaultMaxUnits() 
 
 Dim Entry As String ' Maximum units specified by user 
 Dim R As Resource ' Resource object used in loop 
 
 Entry = InputBox$("Enter the default maximum units for each resource.") 
 
 If IsNumeric(Entry) Then 
 For Each R In ActiveProject.Resources 
 R.MaxUnits = Entry 
 Next R 
 Else 
 MsgBox ("You didn't enter a numeric value.") 
 End If 
 
End Sub
```


