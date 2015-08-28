
# Application.AddSiteColumn Method (Project)
Adds a column that becomes available for the tasks list in the synchronized SharePoint site for the active project.

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AddSiteColumn**(ProjectField, SharePointName)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ProjectField|Optional| ** [PjField](f0df0929-921c-1f33-ab42-192efdaeb64d.md)**|Can be one of a subset of constants in the  **PjField** enumeration, which specifies the project field to show in the new column; cannot be one of the prohibited fields (see Remarks).|
|SharePointName|Optional| **Variant**|The name of the new column.|
|ProjectField|Optional|PJFIELD||
|SharePointName|Optional|VARIANT||

### Return value

 **Boolean**

 **True** if the column is added.


## Remarks
<a name="sectionSection1"> </a>

The  **AddSiteColumn** method gives run-time error 1004, "Application-defined or object-defined error," in the following cases:


- The active project is not associated with a synchronized SharePoint tasks list. If the project is local, you can use the  **LinkToTaskList** method to create a synchronized SharePoint tasks list.
    
- The column name already exists in the SharePoint tasks list. To see the list of column names, open the tasks list in SharePoint, and then choose  **Modify View** on the **LIST** tab. The **Settings - Edit View** page shows all of the column names that are available for the tasks list.
    
- The ProjectField value is a non-task field, such as **pjResourceActualCost**. A SharePoint tasks list shows task fields, not resource fields.
    
- The ProjectField value is an enterprise custom field such as **pjTaskEnterpriseProjectText1**, or a lookup table field such as  **pjTaskResourceEnterpriseRBS**. Local task custom fields, such as  **pjTaskText1**, are valid.
    
- The ProjectField value is one of the prohibited fields in Table 1. These fields are in addition to the prohibited resource fields and enterprise custom fields. The fields are prohibited because they are related to other prohibited fields, or have value types that are not supported in a default SharePoint tasks list.
    
    Although Table 1 and the other prohibited fields may appear to be a long list, of the 1,338 constants in  **PjField**, there are 357â€”including local task custom fieldsâ€”that can be used with the  **AddSiteColumn** method.
    

    **Table 1. Additional prohibited fields**


| **pjTaskActive**| **pjTaskActualOvertimeWork**| **pjTaskACWP**|
| **pjTaskAssignmentDelay**| **pjTaskAssignmentPeakUnits**| **pjTaskAssignmentUnits**|
| **pjTaskBaseline[1-10]BudgetCost**| **pjTaskBaseline[1-10]BudgetWork**| **pjTaskBaseline[1-10]FixedCostAccrual**|
| **pjTaskBaselineBudgetCost**| **pjTaskBaselineBudgetWork**| **pjTaskBaselineFixedCostAccrual**|
| **pjTaskBudgetCost**| **pjTaskBudgetWork**| **pjTaskCalendarGuid**|
| **pjTaskConstraintDate**| **pjTaskConstraintType**| **pjTaskCostRateTable**|
| **pjTaskDeliverableGuid**| **pjTaskDeliverableType**| **pjTaskDemandedRequested**|
| **pjTaskEarnedValueMethod**| **pjTaskEnterpriseOutlineCode[1-30]**| **pjTaskExternalTask**|
| **pjTaskFinishSlack**| **pjTaskFixedCostAccrual**| **pjTaskFreeSlack**|
| **pjTaskGuid**| **pjTaskHideBar**| **pjTaskHyperlink**|
| **pjTaskHyperlinkAddress**| **pjTaskHyperlinkHref**| **pjTaskHyperlinkScreenTip**|
| **pjTaskHyperlinkSubAddress**| **pjTaskID**| **pjTaskIgnoreWarnings**|
| **pjTaskIndicators**| **pjTaskIsAssignment**| **pjTaskLevelAssignments**|
| **pjTaskLevelDelay**| **pjTaskLinkedFields**| **pjTaskManual**|
| **pjTaskMilestone**| **pjTaskNotes**| **pjTaskObjects**|
| **pjTaskOutlineCode[1-10]**| **pjTaskOutlineLevel**| **pjTaskOutlineNumber**|
| **pjTaskPathDrivenSuccessor**| **pjTaskPathDrivingPredecessor**| **pjTaskPathPredecessor**|
| **pjTaskPathSuccessor**| **pjTaskPreleveledFinish**| **pjTaskPreleveledStart**|
| **pjTaskPriority**| **pjTaskResourceType**| **pjTaskStartSlack**|
| **pjTaskStatus**| **pjTaskStatusIndicator**| **pjTaskSubproject**|
| **pjTaskSubprojectReadOnly**| **pjTaskTotalSlack**| **pjTaskType**|
| **pjTaskWarning**| **pjTaskWorkContour**| ****|
It is possible to add a field that already exists in the tasks list, if you use a unique name for the SharePointName parameterâ€”although the value of doing so is questionable.


## Example
<a name="sectionSection2"> </a>

To use the  **AddDurationColumns** macro, create a tasks list in a SharePoint site, create a project in Project Professional, and then use the **LinkToTaskList** method to import the task list. Set a baseline for the active project by using the **Set Baseline** command on the **PROJECT** tab of the ribbon, and then change the duration of some tasks.

The  **AddDurationColumns** macro adds task duration and baseline duration to the list of available columns in the SharePoint task list (see Figure 1).


 **Note**  After you run the  **AddDurationColumns** macro, you must save the project in Project Professional to synchronize your changes with the SharePoint task list.




```
Sub AddDurationColumns()
    Dim success As Boolean
    Dim results As String
    Dim columnName As String
    Dim fieldName As PjField
    results = ""
    
    ' Add the first column.
    fieldName = pjTaskBaselineDurationText
    columnName = "Baseline duration"
    
    ' If the field name exists in the SharePoint tasks list, or fieldName
    ' is one of the prohibited fields, the AddSiteColumn method
    ' returns error 1100.
    On Error Resume Next
    
    success = AddSiteColumn(fieldName, columnName)
    
    If success Then
        results = "Added site column: " &amp; columnName
    Else
        results = "Error in AddSiteColumn: " &amp; columnName
    End If
    
    ' Add the second column.
    fieldName = pjTaskDurationText
    columnName = "Current duration"
    
    success = AddSiteColumn(fieldName, columnName)
    
    If success Then
        results = results &amp; vbCrLf &amp; "Added site column: " &amp; columnName
    Else
        results = results &amp; vbCrLf &amp; "Error in AddSiteColumn: " &amp; columnName
    End If
    
    Debug.Print results
End Sub
```

After you save the project, go to the task list in SharePoint. On the  **LIST** tab, select the **Modify View** command. On the Settings - Edit View page, select the **Baseline duration** field and the **Current duration** field that the **AddDurationColumns** macro added. Figure 1 shows the task list with the two new fields.


**Figure 1. Adding fields to a synchronized SharePoint task list**

![Adding fields to a synchronized tasks list](../images/pj15_VBA_AddSiteColumn.gif)


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Application Object](8eb91712-7784-a102-38c0-19bb056c27e9.md)
 [PjField Enumeration](f0df0929-921c-1f33-ab42-192efdaeb64d.md)
#### Other resources


 [LinkToTaskList Method](65ae7bd0-446f-74dd-15fc-0a260342be90.md)
