
# Pane.View Method (Project)

Returns the active  **View** object.


## Syntax

 _expression_. **View**

 _expression_A variable that represents a  **Pane** object.


### Return Value

 **View**


## Example

The following statement prints the name of the view in the  **Immediate** window in the VBE. For example, if the Team Planner view is active, the statement prints "Team Plannner".


```
Debug.Print ActiveWindow.ActivePane.View.Name
```

