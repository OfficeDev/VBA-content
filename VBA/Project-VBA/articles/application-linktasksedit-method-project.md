---
title: Application.LinkTasksEdit Method (Project)
keywords: vbapj.chm2052
f1_keywords:
- vbapj.chm2052
ms.prod: project-server
api_name:
- Project.Application.LinkTasksEdit
ms.assetid: 51c1d75e-afb6-ae8c-162d-15e24c81bd06
ms.date: 06/08/2017
---


# Application.LinkTasksEdit Method (Project)

Edits task dependencies (task links).


## Syntax

 _expression_. **LinkTasksEdit**( ** _From_**, ** _To_**, ** _Delete_**, ** _Type_**, ** _Lag_**, ** _PredecessorProjectName_**, ** _SuccessorProjectName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _From_|Required|**Long**|**Long**. The identification number of a predecessor task.|
| _To_|Required|**Long**|**Long**. The identification number of a successor task.|
| _Delete_|Optional|**Boolean**|**True** if Project deletes the referenced link. The default value is **False**.|
| _Type_|Optional|**Long**|The relationship between tasks that become linked. Can be one of the [PjTaskLinkType](pjtasklinktype-enumeration-project.md) constants. The default value is **pjFinishToStart**.|
| _Lag_|Optional|**Variant**|The duration between linked tasks in default units. To specify lead time between tasks, use a negative value.|
| _PredecessorProjectName_|Optional|**String**|The name of the subproject in a consolidated project that contains the task identified with  **From**. If **PredecessorProjectName** is omitted, the current project is assumed.|
| _SuccessorProjectName_|Optional|**String**|The name of the subproject in a consolidated project that contains the task identified with  **To**. If **SuccessorProjectName** is omitted, the current project is assumed.|

### Return Value

 **Boolean**


## Example

The following example prompts the user for a range of task identification numbers, and then links the tasks in the range from finish to start. This example assumes the ID range is valid, as well as the absence of any duplicate tasks, null tasks, consolidated projects, and so on.


```vb
Sub LinkFinishToStart() 
 
    Dim FirstID As String ' The ID number of the first task 
    Dim LastID As String ' The ID number of the last task 
    Dim NextID As Long ' The ID number of the next task to link 
 
    FirstID = InputBox$("Enter the ID number of the first task to link:") 
    If FirstID = Empty Then Exit Sub 

    LastID = InputBox$("Enter the ID number of the last task to link:") 
    If LastID = Empty Then Exit Sub 
 
    ' Convert FirstID from String to Long, then "seed" the loop. 
    NextID = CLng(FirstID) 
 
    Do Until NextID = CLng(LastID) 
        LinkTasksEdit From:=NextID, To:=NextID + 1, Type:=pjFinishToStart 
        NextID = NextID + 1 
    Loop 
End Sub
```


