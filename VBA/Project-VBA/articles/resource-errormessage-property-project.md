---
title: Resource.ErrorMessage Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.ErrorMessage
ms.assetid: cb78732f-8c9c-df97-b6bc-c4eb52f4bf16
ms.date: 06/08/2017
---


# Resource.ErrorMessage Property (Project)

Gets errors reported by the  **Import Resources Wizard** and by local resource error checks. Read-only **String**.


## Syntax

 _expression_. **ErrorMessage**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The  **ErrorMessage** property is used by the **Import Resources Wizard** while saving the enterprise resource pool and when **[CheckResourceErrors](application-checkresourceerrors-method-project.md)** and **[EnterpriseResourcesImport](application-enterpriseresourcesimportex-method-project.md)** methods are used.


