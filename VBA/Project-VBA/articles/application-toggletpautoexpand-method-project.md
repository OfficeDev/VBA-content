---
title: Application.ToggleTPAutoExpand Method (Project)
keywords: vbapj.chm1502
f1_keywords:
- vbapj.chm1502
ms.prod: project-server
api_name:
- Project.Application.ToggleTPAutoExpand
ms.assetid: 17520aa8-b364-22be-cdc3-62850e77a228
ms.date: 06/08/2017
---


# Application.ToggleTPAutoExpand Method (Project)

Expands or collapses resource rows in the Team Planner view, where there is more than one assignment within the same time span for a resource.


## Syntax

 _expression_. **ToggleTPAutoExpand**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

You can manually expand or collapse the list of tasks for a single resource by choosing the  **+** or **-** icon next to the resource name, or by using the **[ToggleTPResourceExpand](application-toggletpresourceexpand-method-project.md)** method. The **ToggleTPAutoExpand** method does the same action for all resources.


 **Note**  The  **+** or **-** icon does not show next to the resource name if there are no overlapping assignments for that resource.

The  **ToggleTPAutoExpand** method corresponds to the **Expand Resource Rows** check box on the **Format** tab under **Team Planner Tools** in the ribbon.


## Example

In the following example, at least one resource has overlapping assignments. The  **ToggleResourceRows** macro switches to the Team Planner view and expands or collapses the rows that have overlapping assignments. When a row is expanded, it is easier to see all of the overlapping assignments.


```vb
Sub ToggleResourceRows() 
    ViewApplyEx Name:="Team Planner" 
 
    ToggleTPAutoExpand 
End Sub
```


