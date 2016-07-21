
# Project.TaskTableList Property (Project)

Gets a  **[List](3934c2e8-d810-6571-9a33-1d41edbab87a.md)** object representing all task tables in the project. Read-only **List**.


## Syntax

 _expression_. **TaskTableList**

 _expression_ A variable that represents a **Project** object.


## Example

The following example lists all the task tables in the active project.


```vb
Sub SeeAllTables() 
 
 Dim Temp As Variant 
 Dim TaskTableNames As String 
 
 For Each Temp In ActiveProject.TaskTableList 
 TaskTableNames = TaskTableNames &; vbCrLf &; Temp 
 Next Temp 
 
 MsgBox TaskTableNames 
 
End Sub
```

