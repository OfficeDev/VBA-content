
# Project.Tasks Property (Project)

Gets a  **[Tasks](bc6bb4a5-95a6-9d1f-3e28-92b9548a544a.md)** collection representing the tasks in the project. Read-only **Tasks**.


## Syntax

 _expression_. **Tasks**

 _expression_ A variable that represents a **Project** object.


## Example

The following example displays the name of every task in the active project.


```vb
Sub TaskNames() 
 
 Dim T As Task, Names As String 
 
 For Each T In ActiveProject.Tasks 
 Names = Names &; T.Name &; vbCrLf 
 Next T 
 
 MsgBox Names 
 
End Sub
```

