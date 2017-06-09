---
title: Application.WBSCodeRenumber Method (Project)
keywords: vbapj.chm629
f1_keywords:
- vbapj.chm629
ms.prod: project-server
api_name:
- Project.Application.WBSCodeRenumber
ms.assetid: c71f6dd3-5ea5-de60-7cd5-09134fa5a278
ms.date: 06/08/2017
---


# Application.WBSCodeRenumber Method (Project)

Renumbers work breakdown structure (WBS) codes for either the active project or selected tasks.


## Syntax

 _expression_. **WBSCodeRenumber**( ** _All_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _All_|Optional|**Boolean**|**True** if all tasks in the active project should be renumbered. **False** if only the selected tasks should be renumbered.|

### Return Value

 **Boolean**


## Remarks

Using the  **WBSCodeRenumber** method without specifying any arguments brings up the **WBS Renumber** dialog box, where you can choose whether to renumber selected tasks or the entire project.


