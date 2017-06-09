---
title: Application.FieldConstantToFieldName Method (Project)
ms.prod: project-server
api_name:
- Project.Application.FieldConstantToFieldName
ms.assetid: b8e55035-64e8-fda5-4ad6-9f5e51a55181
ms.date: 06/08/2017
---


# Application.FieldConstantToFieldName Method (Project)

Returns a custom field name for the specified field constant.


## Syntax

 _expression_. **FieldConstantToFieldName**( ** _Field_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required|**Long**|The numerical constant for the custom field. Can be one of the  **[PjField](pjfield-enumeration-project.md)** constants for local custom fields or another **Long** value for enterprise custom fields.|

### Return Value

 **String**


## Remarks

If the Field argument is a local custom field, you can use one of the  **[PjField](pjfield-enumeration-project.md)** constants. If Field is an enterprise custom field, it does not match a **PjField** constant because there can be an unlimited number of enterprise custom fields.


 **Note**  For usability and performance reasons, the number of enterprise custom fields should be limited to a few hundred or less.

You can access project custom fields through the  **ProjectSummaryTask** property.


## Example

The following example shows the difference between the  **[FieldNameToFieldConstant](application-fieldnametofieldconstant-method-project.md)** method and the **FieldConstantToFieldName** method:


1. To use the example, use Project Web App to create an enterprise project text custom field named  **TestEntProjText**. 
    
2. Restart Project with a Project Server profile, so that it includes the new custom field.
    
3. Create a project with some value for the  **TestEntProjText** field, by using the **Project Information** dialog box.
    
4. The  **TestEnterpriseProjectCF** macro uses the **FieldNameToFieldConstant** method to find the projectField number, for example, 190873618.
    
5. The macro shows the number and text value in a message box, by using the  **GetField** method.
    
6. The macro gets the field name by using the  **FieldConstantToFieldName** method, sets a new value by using the **SetField** method, and then shows the field name and new value in another message box.
    





```vb
Sub TestEnterpriseProjectCF() 
    Dim projectField As Long 
    Dim projectFieldName As String 
    Dim message As String 
 
    projectField = FieldNameToFieldConstant("TestEntProjText", pjProject) 
 
    ' Show the enterprise project field number and old value. 
    message = "Enterprise project field number: " &; projectField &; vbCrLf 
    MsgBox message &; ActiveProject.ProjectSummaryTask.GetField(projectField) 
 
    ActiveProject.ProjectSummaryTask.SetField FieldID:=projectField, Value:="This is a new value." 
 
    ' For a demonstration, show the field name from the field number, and verify the new value. 
    projectFieldName = FieldConstantToFieldName(projectField) 
    message = "New value for field: " &; projectFieldName &; vbCrLf 
    MsgBox message &; ActiveProject.ProjectSummaryTask.GetField(projectField) 
End Sub
```

The following example shows the difference in names between the  **pjTaskStart**, **pjTaskStartText**, and similar task fields.


 **Note**  The  **pjTask*Text** fields, such as **pjTaskStartText**, are new in Project. Those fields are used to get data for dates of both automatically and manually scheduled tasks. For example, the **Start** column in a Gantt chart contains **String** data for dates, not **Variant** data. You can use fields such as **pjTaskDuration** in custom field formulas, but not in column headings.

Columns in task views for  **Start**,  **Finish**,  **Duration**, and so forth, contain  **String** data for both auto-scheduled and manually scheduled tasks. The **Duration** column can only use **String** data, so there is no column heading for **pjTaskDuration**.




```vb
Sub TryNewTaskConstants() 
      ' The pj*Text fields return data for the date columns of automatically and manually scheduled tasks. 
    ' For example, FieldConstantToFieldName(pjTaskStartText) returns the column name for Start date strings. 
 
    Debug.Print "pjTaskStart returns: " &; FieldConstantToFieldName(pjTaskStart) 
    Debug.Print "pjTaskStartText returns: " &; FieldConstantToFieldName(pjTaskStartText) _ 
        &; vbCrLf 
 
    Debug.Print "pjTaskFinish returns: " &; FieldConstantToFieldName(pjTaskFinish) 
    Debug.Print "pjTaskFinishText returns: " &; FieldConstantToFieldName(pjTaskFinishText) _ 
        &; vbCrLf 
 
    Debug.Print "pjTaskDuration returns: " &; FieldConstantToFieldName(pjTaskDuration) 
    Debug.Print "pjTaskDurationText returns: " &; FieldConstantToFieldName(pjTaskDurationText) _ 
        &; vbCrLf 
 
    Debug.Print "pjTaskBaselineStart returns: " &; FieldConstantToFieldName(pjTaskBaselineStart) 
    Debug.Print "pjTaskBaselineStartText returns: " &; FieldConstantToFieldName(pjTaskBaselineStartText) 
End Sub
```


