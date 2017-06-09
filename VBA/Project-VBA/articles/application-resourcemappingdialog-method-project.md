---
title: Application.ResourceMappingDialog Method (Project)
keywords: vbapj.chm2255
f1_keywords:
- vbapj.chm2255
ms.prod: project-server
api_name:
- Project.Application.ResourceMappingDialog
ms.assetid: b465a823-769f-7e3e-2f2c-98bda2502e0a
ms.date: 06/08/2017
---


# Application.ResourceMappingDialog Method (Project)

Displays the  **Map Project Resources onto Enterprise Resources** dialog box, for importing local resources to Project Server.


## Syntax

 _expression_. **ResourceMappingDialog**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **Boolean**


## Remarks

To use the  **ResourceMappingDialog** method, a local project must be open and active. If an enterprise project is active, using the **ResourceMappingDialog** method results in the run-time error 1100.

You can use  **ResourceMappingDialog** to avoid the extra step of opening a project with the **[EnterpriseResourcesImportEx](application-enterpriseresourcesimportex-method-project.md)** method or by using the **Import Resources to Enterprise** command on the **Add Resources** drop-down menu of the **Resource** tab in the Ribbon.

 The **ResourceMappingDialog** method is available only in Project Professional.


