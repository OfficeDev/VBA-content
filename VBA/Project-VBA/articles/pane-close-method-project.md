---
title: Pane.Close Method (Project)
ms.prod: project-server
api_name:
- Project.Pane.Close
ms.assetid: 9bd722fd-cd92-9d59-7cdb-9aa40911120a
ms.date: 06/08/2017
---


# Pane.Close Method (Project)

Closes the bottom pane of a window.


## Syntax

 _expression_. **Close**

 _expression_ An expression that returns a **Pane** object.


### Return Value

 **Nothing**


## Example

The following commands exercise the  **Close** method for **Pane** objects.


1. Create and apply a combination view named  **Combo View**.
    
2. Close the bottom pane.
    
3. Apply the  **Combo View** again to open the bottom pane.
    
4. Activate the top pane.
    
5. The  `ActivePane.Close` command does nothing, because the bottom pane is not active.
    
6. Activate the bottom pane.
    
7. Close the bottom pane with the  `ActivePane.Close` command.
    





```vb
ActiveProject.ViewsCombination.Add(Name:="Combo View", TopView:="Gantt Chart", BottomView:="Resource Sheet").Apply 
ActiveWindow.BottomPane.Close 
ActiveProject.ViewsCombination("Combo View").Apply 
ActiveWindow.TopPane.Activate 
ActiveWindow.ActivePane.Close 
ActiveWindow.BottomPane.Activate 
ActiveWindow.ActivePane.Close
```


