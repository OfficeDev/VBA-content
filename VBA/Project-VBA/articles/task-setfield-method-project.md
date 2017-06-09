---
title: Task.SetField Method (Project)
ms.prod: project-server
api_name:
- Project.Task.SetField
ms.assetid: f25de144-79f0-9c19-500a-94708a1b29a1
ms.date: 06/08/2017
---


# Task.SetField Method (Project)

Sets the value of the specified task custom field.


## Syntax

 _expression_. **SetField**( ** _FieldID_**, ** _Value_** )

 _expression_ A variable that represents a **Task** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|For a local custom field, can be one of the  **[PjField](pjfield-enumeration-project.md)** constants for task custom fields. For an enterprise custom field, use the **[FieldNameToFieldConstant](application-fieldnametofieldconstant-method-project.md)** method to get the _FieldID_.|
| _Value_|Required|**String**|The value of the field.|

## Remarks

You can access project custom fields through the  **ProjectSummaryTask** property.


## Example

The following example uses the  **SetField** method and the **GetField** method together with the **FieldNameToFieldConstant** and **FieldConstantToFieldName** methods.


1. To use the example, use Project Web App to create an enterprise project text custom field named  **TestEntProjText**.
    
2. Restart Project Professional with a Project Server profile, so that it includes the new custom field.
    
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
 
    ' For a demonstration, get the field name from the field number, and verify the new value. 
    projectFieldName = FieldConstantToFieldName(projectField) 
    message = "New value for field: " &; projectFieldName &; vbCrLf 
    MsgBox message &; ActiveProject.ProjectSummaryTask.GetField(projectField) 
End Sub
```


