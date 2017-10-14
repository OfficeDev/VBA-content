---
title: Application.EnterpriseResSubstitutionWizard Method (Project)
keywords: vbapj.chm2123
f1_keywords:
- vbapj.chm2123
ms.prod: project-server
api_name:
- Project.Application.EnterpriseResSubstitutionWizard
ms.assetid: 627b04ad-0088-5032-4f05-b6dc8cabe436
ms.date: 06/08/2017
---


# Application.EnterpriseResSubstitutionWizard Method (Project)

Runs the  **Resource Substitution Wizard**. Available in Project Professional only.


## Syntax

 _expression_. **EnterpriseResSubstitutionWizard**( ** _ProjectList_**, ** _PoolOption_**, ** _RBSorResourceList_**, ** _FreezeHorizonDate_**, ** _UpdateProjects_**, ** _SaveReport_**, ** _Path_**, ** _AssignProposedResources_**, ** _LevelProposedBookings_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProjectList_|Optional|**String**|A comma-separated list of portfolio projects in the database.|
| _PoolOption_|Optional|**Long**|Specifies the resource pool option. Can be one of the following  **PjResSubstitutionPoolOption** constants: **pjResSubstitutionResInList**, **pjResSubstitutionResInProject**, or **pjResSubstitutionResInRBS**. The default value is **pjResSubstitutionResInProject**.|
| _RBSorResourceList_|Optional|**String**|The name of the RBS (resource breakdown structure) level to use if  **pjResSubstitutionResInRBS** was specified in the PoolOption argument. If **pjResSubstitutionResInList** was specified in the PoolOption argument, the **RBSorResourceList** argument specifies a comma-separated list of resource names to use.|
| _FreezeHorizonDate_|Optional|**String**|The date of the resource freeze horizon.|
| _UpdateProjects_|Optional|**Boolean**|**True** if the **Resource Substitution Wizard** updates projects with the new resource information. The default value is **True**.|
| _SaveReport_|Optional|**Boolean**|**True** if the **Resource Substitution Wizard** saves a report. The default value is **False**.|
| _Path_|Optional|**String**|The directory to use when creating the report. The default value is the My Documents folder of the current user.|
| _AssignProposedResources_|Optional|**Variant**||
| _LevelProposedBookings_|Optional|**Variant**||

### Return Value

 **Boolean**


## Remarks

No events are fired when using the  **EnterpriseResSubstitutionWizard** method.

The  **EnterpriseResSubstitutionWizard** method does not include a parameter for specifying that resources from the enterprise resource pool should be used.


