---
title: SharedWorkspace.Links Property (Office)
keywords: vbaof11.chm276006
f1_keywords:
- vbaof11.chm276006
ms.prod: office
api_name:
- Office.SharedWorkspace.Links
ms.assetid: 7389c657-8028-3914-cb03-5f2f50c448b5
ms.date: 06/08/2017
---


# SharedWorkspace.Links Property (Office)

Gets a  **[SharedWorkspaceLinks](sharedworkspacelinks-object-office.md)** collection that represents the list of links saved in the current shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Links**

 _expression_ A variable that represents a **SharedWorkspace** object.


## Example

The following example lists the links saved in the current shared workspace.


```
  Dim swsLinks As Office.SharedWorkspaceLinks 
    Set swsLinks = ActiveWorkbook.SharedWorkspace.Links 
    MsgBox "There are " &amp; swsLinks.Count &amp; _ 
        " link(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsLinks = Nothing 

```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

