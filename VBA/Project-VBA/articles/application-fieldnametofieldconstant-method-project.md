---
title: Application.FieldNameToFieldConstant Method (Project)
keywords: vbapj.chm131217
f1_keywords:
- vbapj.chm131217
ms.prod: project-server
api_name:
- Project.Application.FieldNameToFieldConstant
ms.assetid: 0830db06-22a7-3ca5-c9ca-f9efbc360767
ms.date: 06/08/2017
---


# Application.FieldNameToFieldConstant Method (Project)

Returns a  **Long** value for a local custom field or an enterprise custom field name.


## Syntax

 _expression_. **FieldNameToFieldConstant**( ** _FieldName_**, ** _FieldType_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldName_|Required|**String**|The name of the local or enterprise custom field.|
| _FieldType_|Optional|**Long**|The type of field. Can be one of the following  **[PjFieldType](pjfieldtype-enumeration-project.md)** constants: **pjProject**, **pjResource**, or **pjTask**. The default value is **pjTask**.|

### Return Value

 **Long**


## Remarks

If the FieldName argument is a local custom field, the returned value can be a  **[PjField](pjfield-enumeration-project.md)** constant. If FieldName is an enterprise custom field, the returned value does not match a **PjField** constant because there can be an unlimited number of enterprise custom fields.


 **Note**  For usability and performance reasons, the number of enterprise custom fields should be limited to a few hundred or less.

You can access project custom fields through the  **ProjectSummaryTask** property.


## Example

The following example shows the difference between the  **FieldNameToFieldConstant** method and the **[FieldConstantToFieldName](application-fieldconstanttofieldname-method-project.md)** method:


1. To use the example, use Project Web App to create an enterprise project text custom field named  **TestEntProjText**. 
    
2. Restart Project with a Project Server profile so that it includes the new custom field.
    
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


