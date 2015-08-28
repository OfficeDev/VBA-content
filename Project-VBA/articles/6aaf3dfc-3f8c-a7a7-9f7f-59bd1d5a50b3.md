
# Task.LinkPredecessors Method (Project)

 **Last modified:** July 28, 2015

Adds one or more predecessors to the task.

## Syntax

 _expression_. **LinkPredecessors**( **_Tasks_**,  **_Link_**,  **_Lag_**)

 _expression_A variable that represents a  **Task** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Tasks|Required| **Object**| The **Task** or **Tasks** object specified becomes a predecessor of the task specified with **expression**.|
|Link|Optional| **Long**| A constant that specifies the relationship between tasks that become linked. Can be one of the [PjTaskLinkType](141a1145-0eb5-3664-4755-394584aec8ac.md) constants. The default value is **pjFinishToStart**.|
|Lag|Optional| **Variant**|A string that specifies the duration of lag time between linked tasks. To specify lead time between tasks, use an expression for  **Lag** that evaluates to a negative value.|

## Example

The following example prompts the user for the name of a task and then makes the task a predecessor of the selected tasks.


```
Sub LinkTasksFromPredecessor() 
    Dim Entry As String   ' Task name entered by user 
    Dim T As Task         ' Task object used in For Each loop 
    Dim I As Long         ' Used in For loop 
    Dim Exists As Boolean ' Whether or not the task exists 
 
    Entry = InputBox$("Enter the name of a task:") 
 
    Exists = False ' Assume task doesn't exist. 
 
    ' Search active project for the specified task. 
    For Each T In ActiveProject.Tasks 
        If T.Name = Entry Then 
            Exists = True 
            ' Make the task a predecessor of the selected tasks. 
            For I = 1 To ActiveSelection.Tasks.Count 
                ActiveSelection.Tasks(I).LinkPredecessors Tasks:=T 
            Next I 
        End If 
    Next T 
 
    ' If task doesn't exist, display an error and quit the procedure. 
    If Not Exists Then 
        MsgBox ("Task not found.") 
        Exit Sub 
    End If 
End Sub
```

