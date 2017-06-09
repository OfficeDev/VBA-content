---
title: Application.OptionsSecurityTab Method (Project)
keywords: vbapj.chm2504
f1_keywords:
- vbapj.chm2504
ms.prod: project-server
api_name:
- Project.Application.OptionsSecurityTab
ms.assetid: f19ecd9c-2507-e437-7780-cf4998b7fd48
ms.date: 06/08/2017
---


# Application.OptionsSecurityTab Method (Project)

Displays a specific tab of the  **Trust Center** dialog box in Project.


## Syntax

 _expression_. **OptionsSecurityTab**( ** _DefaultTab_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DefaultTab_|Optional|**PjOptionsSecurityTab**|Specifies the tab to open in the  **Trust Center** dialog box. Can be one of the **[PjOptionsSecurityTab](pjoptionssecuritytab-enumeration-project.md)** constants. The default is **pjOptionsSecurityTabPublishers** for the **Trusted Publishers** tab.|

### Return Value

 **Boolean**


