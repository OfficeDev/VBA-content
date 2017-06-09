---
title: SharedWorkspaceTask.Priority Property (Office)
keywords: vbaof11.chm264004
f1_keywords:
- vbaof11.chm264004
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.Priority
ms.assetid: 8e0224a3-9c0c-5c0f-92e8-d7b945236886
ms.date: 06/08/2017
---


# SharedWorkspaceTask.Priority Property (Office)

Gets or sets the status of the specified shared workspace task. Read/write.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Priority**

 _expression_ Required. A variable that represents a **[SharedWorkspaceTask](sharedworkspacetask-object-office.md)** object.


## Remarks

The shared workspace task schema on the server can be customized. Customization of the schema may affect the task priority enumeration when the  **Add** or **Save** method is called. **Priority** property values are mapped as follows:




- Downloaded value 1 is mapped to  **msoSharedWorkspaceTaskPriority** 1 ( **msoSharedWorkspaceTaskPriorityHigh** ). Downloaded values 2 through N-1 are mapped to **msoSharedWorkspaceTaskPriority** 2 ( **msoSharedWorkspaceTaskPriorityNormal** ). Downloaded value N is mapped to **msoSharedWorkspaceTaskPriority** 3 ( **msoSharedWorkspaceTaskPriorityLow** ).
    
- Uploaded enumeration values 1 through 3 are mapped to schema values 1 through 3. If a user-specified value does not map to any value defined in the schema, the user-specified value is silently ignored and the  **Status** property is not updated on the server.
    



## See also


#### Concepts


[SharedWorkspaceTask Object](sharedworkspacetask-object-office.md)
#### Other resources


[SharedWorkspaceTask Object Members](sharedworkspacetask-members-office.md)

