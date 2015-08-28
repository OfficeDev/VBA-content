
# Application.UpdateTasks Method (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Updates the selected tasks.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **UpdateTasks**( **_PercentComplete_**,  **_ActualDuration_**,  **_RemainingDuration_**,  **_ActualStart_**,  **_ActualFinish_**,  **_Notes_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PercentComplete|Optional| **Variant**|The percent complete of the active tasks.|
|ActualDuration|Optional| **Variant**|The actual duration of the selected tasks.|
|RemainingDuration|Optional| **Variant**|The remaining duration of the selected tasks.|
|ActualStart|Optional| **Variant**|The actual start date of the selected tasks.|
|ActualFinish|Optional| **Variant**|The actual finish date of the selected tasks.|
|Notes|Optional| **String**|Comments in the Notes field for the selected tasks. The value can be text only, not Rich Text Format (RTF) as in the  **Notes** dialog box.|

### Return Value

 **Boolean**


## Remarks
<a name="sectionSection1"> </a>

Using the  **UpdateTasks** method without specifying any arguments displays the **Update Tasks** dialog box.


## Example
<a name="sectionSection2"> </a>

The following example creates a task named "TestTask-1", updates the task to 50% complete, and then deletes the task. 


```
Sub Update_Tasks() 
 
 'Activate Gantt Chart 
 ViewApply Name:="Gantt Chart" 
 
 'Create a task 
 RowInsert 
 SetTaskField Field:="Name", Value:="TestTask-1" 
 SetTaskField Field:="Duration", Value:="2" 
 
 'Update the percent complete of the new task. 
 UpdateTasks PercentComplete:="50" 
 
 'Delete the new task 
 ActiveProject.Tasks("TestTask-1").Delete 
End Sub
```

