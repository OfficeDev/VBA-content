---
title: OutlineCode.FieldID Property (Project)
ms.prod: project-server
api_name:
- Project.OutlineCode.FieldID
ms.assetid: eea0a697-08f9-c4f5-358a-6b90bd08271e
ms.date: 06/08/2017
---


# OutlineCode.FieldID Property (Project)

Gets the identification number of the local outline code. Read-only  **PjCustomField**.


## Syntax

 _expression_. **FieldID**

 _expression_ A variable that represents an **OutlineCode** object.


## Remarks

To get the ID of an enterprise text custom field, use the  **[FieldNameToFieldConstant](application-fieldnametofieldconstant-method-project.md)** method.


 **Note**  In Office Project 2007 and later versions, the enterprise constants in  **PjCustomField** do not apply. Project Server can have an unlimited number of enterprise text custom fields that use a hierarchical lookup table. For usability and performance reasons, the number of enterprise custom fields should be limited to a few hundred or less.

You can access project outline codes and custom fields through the project summary task, which is  `Task(0)`. For a task outline code, the  **FieldID** can be one of the following **[PjCustomField](pjcustomfield-enumeration-project.md)** constants:


||
|:-----|
|**pjCustomTaskOutlineCode1**|
|**pjCustomTaskOutlineCode2**|
|**pjCustomTaskOutlineCode3**|
|**pjCustomTaskOutlineCode4**|
|**pjCustomTaskOutlineCode5**|
|**pjCustomTaskOutlineCode6**|
|**pjCustomTaskOutlineCode7**|
|**pjCustomTaskOutlineCode8**|
|**pjCustomTaskOutlineCode9**|
|**pjCustomTaskOutlineCode10**|
For a resource outline code, the  **FieldID** can be one of the following **PjCustomField** constants:


||
|:-----|
|**pjCustomResourceOutlineCode1**|
|**pjCustomResourceOutlineCode2**|
|**pjCustomResourceOutlineCode3**|
|**pjCustomResourceOutlineCode4**|
|**pjCustomResourceOutlineCode5**|
|**pjCustomResourceOutlineCode6**|
|**pjCustomResourceOutlineCode7**|
|**pjCustomResourceOutlineCode8**|
|**pjCustomResourceOutlineCode9**|
|**pjCustomResourceOutlineCode10**|

