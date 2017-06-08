---
title: SharedWorkspaceLinks.Creator Property (Office)
ms.prod: office
api_name:
- Office.SharedWorkspaceLinks.Creator
ms.assetid: 857d2a06-1572-19b7-ebae-6532937adad4
ms.date: 06/08/2017
---


# SharedWorkspaceLinks.Creator Property (Office)

Gets a 32-bit integer that indicates the application in which the  **SharedWorkspaceLinks** object was created. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Creator**

 _expression_ A variable that represents a **SharedWorkspaceLinks** object.


### Return Value

Long


## Remarks

As an example, if the object was created in Microsoft Word, this property returns 1297307460, which represents the string "MSWD"; in Microsoft Excel, this property returns 1480803660. This value can also be represented by the constant wdCreatorCode in Word, or xlCreatorCode in Excel. The  **Creator** property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.

The  **Creator** property always returns the numeric identifier for the active application, just as the **Application** property always returns the name of the active applicatin in string form. Use the **CreatedBy** property of the **SharedWorkspaceLink** objects to return the name of the individual who created the object. Use document properties to return information about the authors of Office documents.


## See also


#### Concepts


[SharedWorkspaceLinks Object](sharedworkspacelinks-object-office.md)
#### Other resources


[SharedWorkspaceLinks Object Members](sharedworkspacelinks-members-office.md)

