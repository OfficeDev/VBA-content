---
title: Application.EditGoTo Method (Project)
keywords: vbapj.chm213
f1_keywords:
- vbapj.chm213
ms.prod: project-server
api_name:
- Project.Application.EditGoTo
ms.assetid: cd2c886b-fddf-d7b8-8f16-51a3af5f0005
ms.date: 06/08/2017
---


# Application.EditGoTo Method (Project)

Scrolls to a resource, task, or date.


## Syntax

 _expression_. **EditGoTo**( ** _ID_**, ** _Date_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ID_|Optional|**Long**|A number that specifies the identification number of the task or resource to display in the active pane.|
| _Date_|Optional|**Variant**|A number or string that specifies the first date to display in the active pane.|

### Return Value

 **Boolean**


## Example

The following example prompts the user for a date or a task name, and then scrolls to that date or task in the active pane. It assumes the user is in a task view.


```vb
Sub PromptUserForEditGotoArguments() 
 
 Dim Entry As String ' Date or task name entered by user 
 
 Entry = InputBox$("Enter a date or a task name to which you want to scroll in the active pane.") 
 
 ' If user enters a date, scroll to a date in the active pane. 
 If IsDate(Entry) Then 
 EditGoTo Date:=Entry 
 ' Otherwise, scroll to a task in the active pane. 
 Else 
 EditGoTo ID:=ActiveProject.Tasks(Entry).ID 
 End If 
 
End Sub
```


