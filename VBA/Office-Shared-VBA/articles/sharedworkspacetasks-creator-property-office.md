---
title: SharedWorkspaceTasks.Creator Property (Office)
ms.prod: office
api_name:
- Office.SharedWorkspaceTasks.Creator
ms.assetid: e89b63e8-6ae4-8f45-615c-eee5f0b6e8ad
ms.date: 06/08/2017
---


# SharedWorkspaceTasks.Creator Property (Office)

Gets a 32-bit integer that indicates the application in which the  **SharedWorkspaceTasks** object was created. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Creator**

 _expression_ A variable that represents a **SharedWorkspaceTasks** object.


### Return Value

Long


## Remarks

As an example, if the object was created in Microsoft Word, this property returns 1297307460, which represents the string "MSWD"; in Microsoft Excel, this property returns 1480803660. This value can also be represented by the constant wdCreatorCode in Word, or xlCreatorCode in Excel. The  **Creator** property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.

The  **Creator** property always returns the numeric identifier for the active application, just as the **Application** property always returns the name of the active applicatin in string form. Use the **CreatedBy** property of the **SharedWorkspaceTask** object to return the name of the individual who created the objects. Use document properties to return information about the authors of Office documents.


## Example

This example displays a message about the creator of "myObject" variable.


```
Set myObject = ActiveDocument 
If myObject.Creator = wdCreatorCode Then 
    MsgBox "This is a Microsoft Word object" 
Else 
    MsgBox "This is not a Microsoft Word object" 
End If 

```


## See also


#### Concepts


[SharedWorkspaceTasks Object](sharedworkspacetasks-object-office.md)
#### Other resources


[SharedWorkspaceTasks Object Members](sharedworkspacetasks-members-office.md)

