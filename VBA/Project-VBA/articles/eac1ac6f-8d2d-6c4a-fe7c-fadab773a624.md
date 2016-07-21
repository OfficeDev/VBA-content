
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

