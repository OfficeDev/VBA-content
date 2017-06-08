---
title: SharedWorkspaceLinks.Add Method (Office)
keywords: vbaof11.chm271003
f1_keywords:
- vbaof11.chm271003
ms.prod: office
api_name:
- Office.SharedWorkspaceLinks.Add
ms.assetid: 76c1fe99-14de-7276-0c5c-fd54f6d0a6ce
ms.date: 06/08/2017
---


# SharedWorkspaceLinks.Add Method (Office)

Adds a link to the list of links in a shared workspace.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Add**( **_URL_**, **_Description_**, **_Notes_** )

 _expression_ Required. A variable that represents a **[SharedWorkspaceLinks](sharedworkspacelinks-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _URL_|Required|**String**|The address of the Web site to which a link is being added.|
| _Description_|Optional|**String**|Description of the link.|
| _Notes_|Optional|**String**|Notes about the link.|

### Return Value

SharedWorkspaceLink


## Example

The following example adds a new link to the links collection of the shared workspace.


```
    Dim swsLink As Office.SharedWorkspaceLink 
    Set swsLink = ActiveWorkbook.SharedWorkspace.Links.Add( _ 
        "http://msdn.microsoft.com", _ 
        "Microsoft Developer Network Home Page", _ 
        "My favorite developer site!") 
    MsgBox "New link: " &amp; swsLink.Description, _ 
        vbInformation + vbOKOnly, _ 
        "New Link in Shared Workspace" 
    Set swsLink = Nothing 

```


## See also


#### Concepts


[SharedWorkspaceLinks Object](sharedworkspacelinks-object-office.md)
#### Other resources


[SharedWorkspaceLinks Object Members](sharedworkspacelinks-members-office.md)

