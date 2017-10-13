---
title: Application.ResourceAssignment Method (Project)
keywords: vbapj.chm212
f1_keywords:
- vbapj.chm212
ms.prod: project-server
api_name:
- Project.Application.ResourceAssignment
ms.assetid: aceb1802-4b5f-0ad3-bd14-ce77c24705fb
ms.date: 06/08/2017
---


# Application.ResourceAssignment Method (Project)

Assigns, removes, or replaces the resources of the selected tasks, or changes the number of units for a resource.

## Syntax

_expression_. **ResourceAssignment** (**_Resources_**, **_Operation_**, **_With_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Resources_|Optional|**String**|The names of the resources to be assigned, removed, or replaced in the selected tasks. <br/><br/>**Note**  Project will not assign a resource if thousands separators or decimal separators are included in the unit values.|
| _Operation_|Optional|**Long**|If _Operation_ is omitted, Project assigns the resources to the selected tasks. The default value is **pjAssign**. Can be one of the **[PjResAssignOperation constants](#pjresassignoperation-constants)**.|
| _With_|Optional|**String**|When used with the **pjReplace** constant for _Operation_, specifies the names of the resources that replace the resources of the selected tasks.|

<br/>

#### PjResAssignOperation constants

|**Constant**|**Description**|
|:-----|:-----|
|**pjAssign**|Assigns the specified resources to the selected tasks.|
|**pjRemove**|Removes the specified resources from the selected tasks.|
|**pjReplace**|The resources specified by _With_ replace the resources specified by _Resources_.|
|**pjChange**|Changes the resource units for the specified resource. This constant can be used only for a single resource.|

<br/>

### Return value

 **Boolean**


## Remarks

You can use the _Resources_ parameter to specify that a resource assignment is requested or demanded when using the Resource Substitution Wizard. For example, the following macro specifies that the assignment of r1 to the selected task is a requested assignment.

```vb
Sub RequestAssignment()
    ResourceAssignment Resources:="r1[100%, R]", Operation:=pjChange, With:="" 
End Sub
```

> [!NOTE]
> When using the _Resources_ parameter in this way, **D** specifies "Demand," **R** specifies "Request," and **N** specifies "None." In addition, spaces are not allowed between the units value and the Request/Demand value. For example, `Resources:="100%,R"` works, but `Resources:="100%, R"` does not.

> The Resource Substitution Wizard cannot substitute material resources. Therefore, you cannot request or demand a material resource for a particular assignment by using the _Resources_ parameter.


## Example

The following example prompts the user for the name of a resource, and then assigns that resource to the selected tasks.


```vb
Sub AssignResourceToSelectedTasks() 
 
    Dim Entry As String     ' The name of the resource to add to selected tasks 
    Dim R As Resource       ' Resource object used in For Each...Next loop 
    Dim Found As Boolean    ' Whether or not the resource is in the active project 
 
    Entry = InputBox$("Enter the name of the resource you want to add to the selected tasks.") 
     
    ' Assume resource doesn't exist in the active project. 
    Found = False 
 
    ' Look for the resource. 
    For Each R In ActiveProject.Resources 
        If Entry = R.Name Then Found = True 
    Next R 
 
    ' If the resource is found, then assign it to selected tasks. 
    If Found Then 
        ResourceAssignment Resources:=Entry, Operation:=pjAssign 
    ' Otherwise, tell user the resource doesn't exist. 
    Else 
        MsgBox ("There is no resource in the active project named " &; Entry &; ".") 
    End If 
     
End Sub
```


