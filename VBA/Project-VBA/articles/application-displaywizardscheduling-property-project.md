---
title: Application.DisplayWizardScheduling Property (Project)
ms.prod: project-server
api_name:
- Project.Application.DisplayWizardScheduling
ms.assetid: abcd5660-1eef-d53b-548f-6ead0c57f836
ms.date: 06/08/2017
---


# Application.DisplayWizardScheduling Property (Project)

 **True** if the **Planning Wizard** displays messages about scheduling problems. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayWizardScheduling**

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


