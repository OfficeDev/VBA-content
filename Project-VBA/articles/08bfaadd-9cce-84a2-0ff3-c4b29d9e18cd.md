
# Project.Tasks Property (Project)

 **Last modified:** July 28, 2015

Gets a  ** [Tasks](bc6bb4a5-95a6-9d1f-3e28-92b9548a544a.md)** collection representing the tasks in the project. Read-only **Tasks**.

## Syntax

 _expression_. **Tasks**

 _expression_A variable that represents a  **Project** object.


## Example

The following example displays the name of every task in the active project.


```
Sub TaskNames() 
 
 Dim T As Task, Names As String 
 
 For Each T In ActiveProject.Tasks 
 Names = Names &amp; T.Name &amp; vbCrLf 
 Next T 
 
 MsgBox Names 
 
End Sub
```

