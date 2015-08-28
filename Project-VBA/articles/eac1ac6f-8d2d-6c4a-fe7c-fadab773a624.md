
# Application.DisplayPlanningWizard Property (Project)

 **Last modified:** July 28, 2015

 **True** if the PlanningWizard is active. Read/write **Boolean**.

## Syntax

 _expression_. **DisplayPlanningWizard**

 _expression_A variable that represents an  **Application** object.


## Example

The following example resets the PlanningWizard to its default settings.


```
Sub ResetWizard() 
 Application.DisplayPlanningWizard = True 
 Application.DisplayWizardErrors = True 
 Application.DisplayWizardScheduling = True 
 Application.DisplayWizardUsage = True 
End Sub
```

