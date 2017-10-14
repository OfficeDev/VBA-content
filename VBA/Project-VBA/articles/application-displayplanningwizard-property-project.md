---
title: Application.DisplayPlanningWizard Property (Project)
ms.prod: project-server
api_name:
- Project.Application.DisplayPlanningWizard
ms.assetid: eac1ac6f-8d2d-6c4a-fe7c-fadab773a624
ms.date: 06/08/2017
---


# Application.DisplayPlanningWizard Property (Project)

 **True** if the PlanningWizard is active. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayPlanningWizard**

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


