---
title: Application.ToggleTPResourceExpand Method (Project)
keywords: vbapj.chm73
f1_keywords:
- vbapj.chm73
ms.prod: project-server
api_name:
- Project.Application.ToggleTPResourceExpand
ms.assetid: a4e39a14-3ba7-25b0-470e-a49c5586d490
ms.date: 06/08/2017
---


# Application.ToggleTPResourceExpand Method (Project)

Expands or collapses the specified resource row in the Team Planner view, when there is more than one assignment within the same time span for the resource.


## Syntax

 _expression_. **ToggleTPResourceExpand**( ** _ResourceUniqueID_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ResourceUniqueID_|Optional|**Long**|Unique identification number of the resource.|

### Return Value

 **Boolean**


## Remarks

You can manually expand or collapse the list of tasks for a single resource, by choosing the  **+** or **?** icon next to the resource name. To do the same action for all resources, use the **[ToggleTPAutoExpand](application-toggletpautoexpand-method-project.md)** method.


 **Note**  The  **+** or **?** icon does not show next to the resource name if there are no overlapping assignments for that resource.

The  **ToggleTPResourceExpand** method corresponds to the **Expand Resource Rows** check box on the **Format** tab under **Team Planner Tools** in the ribbon, but affects only the specified resource.


## Example

In the following example, the R2 resource has overlapping assignments. The  **ToggleTheResourceRow** macro switches to the Team Planner view and expands or collapses the row for R2. When the row is expanded, it is easier to see all of the overlapping assignments.


```vb
Sub ToggleTheResourceRow() 
    Dim resourceUid As Long 
 
    ViewApplyEx Name:="Team Planner" 
 
    resourceUid = ActiveProject.Resources("R2").UniqueID 
    ToggleTPResourceExpand (resourceUid) 
End Sub
```


