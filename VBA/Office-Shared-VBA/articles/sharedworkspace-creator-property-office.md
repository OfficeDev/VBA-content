---
title: SharedWorkspace.Creator Property (Office)
ms.prod: office
api_name:
- Office.SharedWorkspace.Creator
ms.assetid: 167fdd22-50ab-9b27-f594-27c38d88a4a9
ms.date: 06/08/2017
---


# SharedWorkspace.Creator Property (Office)

Gets a 32-bit integer that indicates the application in which the  **SharedWorkspace** object was created. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Creator**

 _expression_ A variable that represents a **SharedWorkspace** object.


### Return Value

Long


## Remarks

As an example, if the object was created in Microsoft Word, this property returns 1297307460, which represents the string "MSWD"; in Microsoft Excel, this property returns 1480803660. This value can also be represented by the constant wdCreatorCode in Word, or xlCreatorCode in Excel. The  **Creator** property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

