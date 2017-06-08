---
title: Application.InformationDialog Method (Project)
keywords: vbapj.chm217
f1_keywords:
- vbapj.chm217
ms.prod: project-server
api_name:
- Project.Application.InformationDialog
ms.assetid: 644b39d6-be73-5a07-4376-02df25d31a02
ms.date: 06/08/2017
---


# Application.InformationDialog Method (Project)

Displays the  **Assignment Information**,  **Resource Information**, or  **Task Information** dialog box for the selected assignment, resource, or task.


## Syntax

 _expression_. **InformationDialog**( ** _Tab_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Tab_|Optional|**Long**|The tab to display in the ** Assignment Information**, **Resource Information**, or  **Task Information** dialog box.|

### Return Value

 **Boolean**


## Remarks

If multiple items are selected, the  **InformationDialog** method displays the **Multiple Assignment Information**,  **Multiple Resource Information**, or  **Multiple Task Information** dialog box.

If an assignment is selected, Tab can be one of the following  **PjInformationTab** constants: **pjAssignmentGeneralTab**, **pjAssignmentTrackingTab**, or **pjAssignmentNotesTab**.

If a resource is selected, Tab can be one of the following  **PjInformationTab** constants: **pjResourceGeneralTab**, **pjResourceWorkingTimeTab**, **pjResourceCostsTab**, or **pjResourceNotesTab**.

If a task is selected, Tab can be one of the following  **PjInformationTab** constants: **pjTaskGeneralTab**, **pjTaskPredecessorsTab**, **pjTaskResourcesTab**, **pjTaskAdvancedTab**, or **pjTaskNotesTab**.


