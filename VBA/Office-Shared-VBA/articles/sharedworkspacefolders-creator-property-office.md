---
title: SharedWorkspaceFolders.Creator Property (Office)
ms.prod: office
api_name:
- Office.SharedWorkspaceFolders.Creator
ms.assetid: 65d21c1c-dac0-ce89-25a0-0fa033e973e5
ms.date: 06/08/2017
---


# SharedWorkspaceFolders.Creator Property (Office)

Gets a 32-bit integer that indicates the application in which the  **SharedWorkspaceFolders** object was created. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Creator**

 _expression_ A variable that represents a **SharedWorkspaceFolders** object.


### Return Value

Long


## Remarks

As an example, if the object was created in Microsoft Word, this property returns 1297307460, which represents the string "MSWD"; in Microsoft Excel, this property returns 1480803660. This value can also be represented by the constant wdCreatorCode in Word, or xlCreatorCode in Excel. The  **Creator** property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.

The  **Creator** property always returns the numeric identifier for the active application, just as the **Application** property always retobjmemName> object to return the name of the individual who created the object. Use document properties to return information about the authors of Office documents.


## See also


#### Concepts


[SharedWorkspaceFolders Object](sharedworkspacefolders-object-office.md)
#### Other resources


[SharedWorkspaceFolders Object Members](sharedworkspacefolders-members-office.md)

