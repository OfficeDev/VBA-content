---
title: Application.DisplayWizardErrors Property (Project)
keywords: vbapj.chm131754
f1_keywords:
- vbapj.chm131754
ms.prod: project-server
api_name:
- Project.Application.DisplayWizardErrors
ms.assetid: b0af54ec-392f-b84d-3dcc-cc52c991b66d
ms.date: 06/08/2017
---


# Application.DisplayWizardErrors Property (Project)

 **True** if the **Planning Wizard** displays messages about errors. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayWizardErrors**

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


