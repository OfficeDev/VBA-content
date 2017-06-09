---
title: SharedWorkspaceLinks Object (Office)
keywords: vbaof11.chm271000
f1_keywords:
- vbaof11.chm271000
ms.prod: office
api_name:
- Office.SharedWorkspaceLinks
ms.assetid: b226b376-9d8c-659a-9551-6341bbebed6f
ms.date: 06/08/2017
---


# SharedWorkspaceLinks Object (Office)

A collection of the  **[SharedWorkspaceLink](sharedworkspacelink-object-office.md)** objects in the current shared workspace.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Links](sharedworkspace-links-property-office.md)** property of the **[SharedWorkspace](sharedworkspace-object-office.md)** object to return a **SharedWorkspaceLinks** collection.


```
    Dim swsLinks As Office.SharedWorkspaceLinks 
    Set swsLinks = ActiveWorkbook.SharedWorkspace.Links 
    MsgBox "There are " &amp; swsLinks.Count &amp; _ 
        " link(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsLinks = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Add](sharedworkspacelinks-add-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](sharedworkspacelinks-application-property-office.md)|
|[Count](sharedworkspacelinks-count-property-office.md)|
|[Creator](sharedworkspacelinks-creator-property-office.md)|
|[Item](sharedworkspacelinks-item-property-office.md)|
|[ItemCountExceeded](sharedworkspacelinks-itemcountexceeded-property-office.md)|
|[Parent](sharedworkspacelinks-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
