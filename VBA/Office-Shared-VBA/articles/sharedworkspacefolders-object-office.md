---
title: SharedWorkspaceFolders Object (Office)
keywords: vbaof11.chm269000
f1_keywords:
- vbaof11.chm269000
ms.prod: office
api_name:
- Office.SharedWorkspaceFolders
ms.assetid: a9020edc-f199-6bab-75d1-c2bdc2a547d3
ms.date: 06/08/2017
---


# SharedWorkspaceFolders Object (Office)

A collection of the  **SharedWorkspaceFolder** objects in the current shared workspace.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **Folders** property of the **SharedWorkspace** object to return a **SharedWorkspaceFolders** collection.


```
    Dim swsFolders As Office.SharedWorkspaceFolders 
    Set swsFolders = ActiveWorkbook.SharedWorkspace.Folders 
    MsgBox "There are " &amp; swsFolders.Count &amp; _ 
        " folder(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFolders = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Add](sharedworkspacefolders-add-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](sharedworkspacefolders-application-property-office.md)|
|[Count](sharedworkspacefolders-count-property-office.md)|
|[Creator](sharedworkspacefolders-creator-property-office.md)|
|[Item](sharedworkspacefolders-item-property-office.md)|
|[ItemCountExceeded](sharedworkspacefolders-itemcountexceeded-property-office.md)|
|[Parent](sharedworkspacefolders-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
