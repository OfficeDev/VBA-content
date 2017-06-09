---
title: Application.EnterpriseResourcesImportEx Method (Project)
keywords: vbapj.chm2090
f1_keywords:
- vbapj.chm2090
ms.prod: project-server
api_name:
- Project.Application.EnterpriseResourcesImportEx
ms.assetid: 58b92ff5-da61-07cc-daca-b56e4270a8a4
ms.date: 06/08/2017
---


# Application.EnterpriseResourcesImportEx Method (Project)

Imports local resources to the enterprise resource pool, or starts the  **Resource Import Wizard**.


## Syntax

 _expression_. **EnterpriseResourcesImportEx**( ** _LocalRUIDs_**, ** _UseImportColumn_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LocalRUIDs_|Optional|**String**|A comma-delimited list of the unique ID numbers of the local resources.|
| _UseImportColumn_|Optional|**Boolean**|**True** if the **EnterpriseResourcesImportEx** method uses the **Import** column.|

### Return Value

 **Boolean**


## Remarks

Using the  **EnterpriseResourcesImportEx** method with no arguments starts the **Resource Import Wizard**. Used that way, the method corresponds to the  **Import Resources to Enterprise** command in the **Add Resources** drop-down menu of the **Resource** tab in the Ribbon.


 **Note**  The  **EnterpriseResourcesImportEx** method starts the file **Open** dialog box, with a list of enterprise projects. You can open an enterprise project that contains local resources or a local project. Alternately, you can use the **[ResourceMappingDialog](application-resourcemappingdialog-method-project.md)** method instead of **EnterpriseResourcesImportEx** to avoid the extra step of opening a project.

The  **EnterpriseResourcesImportEx** method is available in Project Professional only and requires a connection with Project Server.


