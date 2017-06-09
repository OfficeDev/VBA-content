---
title: SharedWorkspaceTasks Object (Office)
keywords: vbaof11.chm265000
f1_keywords:
- vbaof11.chm265000
ms.prod: office
api_name:
- Office.SharedWorkspaceTasks
ms.assetid: de26341f-44d1-131e-1dbe-e31f3f68e312
ms.date: 06/08/2017
---


# SharedWorkspaceTasks Object (Office)

A collection of the  **[SharedWorkspaceTask](sharedworkspacetask-object-office.md)** objects in the current shared workspace site.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Tasks](sharedworkspace-tasks-property-office.md)** property of the **[SharedWorkspace](sharedworkspace-object-office.md)** object to return a **SharedWorkspaceTasks** collection.


```
    Dim swsTasks As Office.SharedWorkspaceTasks 
    Set swsTasks = ActiveWorkbook.SharedWorkspace.Tasks 
    MsgBox "There are " &amp; swsTasks.Count &amp; _ 
        " task(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsTasks = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Add](sharedworkspacetasks-add-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](sharedworkspacetasks-application-property-office.md)|
|[Count](sharedworkspacetasks-count-property-office.md)|
|[Creator](sharedworkspacetasks-creator-property-office.md)|
|[Item](sharedworkspacetasks-item-property-office.md)|
|[ItemCountExceeded](sharedworkspacetasks-itemcountexceeded-property-office.md)|
|[Parent](sharedworkspacetasks-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
