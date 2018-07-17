---
title: SharedWorkspaceMembers Object (Office)
keywords: vbaof11.chm273000
f1_keywords:
- vbaof11.chm273000
ms.prod: office
api_name:
- Office.SharedWorkspaceMembers
ms.assetid: 2d0e6ce0-79ef-3030-b1af-465428314b15
ms.date: 06/08/2017
---


# SharedWorkspaceMembers Object (Office)

A collection of the  **[SharedWorkspaceMember](sharedworkspacemember-object-office.md)** objects in the current shared workspace site.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Members](sharedworkspace-members-property-office.md)** property of the **[SharedWorkspace](sharedworkspace-object-office.md)** object to return a **SharedWorkspaceMembers** collection.


```
    Dim swsMembers As Office.SharedWorkspaceMembers 
    Set swsMembers = ActiveWorkbook.SharedWorkspace.Members 
    MsgBox "There are " &amp; swsMembers.Count &amp; _ 
        " member(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsMembers = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Add](sharedworkspacemembers-add-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](sharedworkspacemembers-application-property-office.md)|
|[Count](sharedworkspacemembers-count-property-office.md)|
|[Creator](sharedworkspacemembers-creator-property-office.md)|
|[Item](sharedworkspacemembers-item-property-office.md)|
|[ItemCountExceeded](sharedworkspacemembers-itemcountexceeded-property-office.md)|
|[Parent](sharedworkspacemembers-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
