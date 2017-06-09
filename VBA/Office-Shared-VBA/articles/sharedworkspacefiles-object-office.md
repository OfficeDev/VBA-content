---
title: SharedWorkspaceFiles Object (Office)
keywords: vbaof11.chm267000
f1_keywords:
- vbaof11.chm267000
ms.prod: office
api_name:
- Office.SharedWorkspaceFiles
ms.assetid: 5e2937f7-f794-dffb-a1ec-69ea9a9e3546
ms.date: 06/08/2017
---


# SharedWorkspaceFiles Object (Office)

A collection of the  **[SharedWorkspaceFile](sharedworkspacefile-object-office.md)** objects in the current shared workspace.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Files](sharedworkspace-files-property-office.md)** property of the **[SharedWorkspace](sharedworkspace-object-office.md)** object to return a **SharedWorkspaceFiles** collection.


```
    Dim swsFiles As Office.SharedWorkspaceFiles 
    Set swsFiles = ActiveWorkbook.SharedWorkspace.Files 
    MsgBox "There are " &amp; swsFiles.Count &amp; _ 
        " file(s) 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFiles = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Add](sharedworkspacefiles-add-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](sharedworkspacefiles-application-property-office.md)|
|[Count](sharedworkspacefiles-count-property-office.md)|
|[Creator](sharedworkspacefiles-creator-property-office.md)|
|[Item](sharedworkspacefiles-item-property-office.md)|
|[ItemCountExceeded](sharedworkspacefiles-itemcountexceeded-property-office.md)|
|[Parent](sharedworkspacefiles-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
