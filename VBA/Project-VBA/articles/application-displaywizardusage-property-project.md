---
title: Application.DisplayWizardUsage Property (Project)
keywords: vbapj.chm131756
f1_keywords:
- vbapj.chm131756
ms.prod: project-server
api_name:
- Project.Application.DisplayWizardUsage
ms.assetid: 3b4362ca-c748-3da8-0e1d-8d0baa1c3d69
ms.date: 06/08/2017
---


# Application.DisplayWizardUsage Property (Project)

 **True** if the **Planning Wizard** displays tips about using Project more effectively. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayWizardUsage**

 _expression_ A variable that represents an **Application** object.


## Example

The following example resets the PlanningWizard to its default settings.


```vb
Sub ResetWizard() 
 Application.DisplayPlanningWizard = True 
 Application.DisplayWizardErrors = True 
 Application.DisplayWizardScheduling = True 
 Application.DisplayWizardUsage = True 
End Sub
```


