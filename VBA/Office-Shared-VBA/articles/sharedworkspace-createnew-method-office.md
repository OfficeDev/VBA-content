---
title: SharedWorkspace.CreateNew Method (Office)
keywords: vbaof11.chm276008
f1_keywords:
- vbaof11.chm276008
ms.prod: office
api_name:
- Office.SharedWorkspace.CreateNew
ms.assetid: 67fbf788-bca0-f83d-acb5-a756bf0ddfb4
ms.date: 06/08/2017
---


# SharedWorkspace.CreateNew Method (Office)

Creates a document workspace site on the server and adds the active document to the new shared workspace site.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **CreateNew**( **_URL_**, **_Name_** )

 _expression_ A variable that represents a **SharedWorkspace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _URL_|Optional|**Variant**|The URL for the parent folder in which the new shared workspace is to be created. If you do not supply a URL, the site is created in the user's default server location.|
| _Name_|Optional|**Variant**| The name of the new shared workspace site. The default value is the name of the active document without its file name extension. For example, if you create a workspace site for "Budget.xls", the name of the new site becomes "Budget".|

## Remarks

Use the  **CreateNew** method to create a shared workspace site for the active document. Omit the 2 optional arguments to create the site using the name of the active document in the user's default server location.

The  **CreateNew** method raises an error if the active document has changes that have not been saved. The document must be saved before you can add it to a shared workspace site.


 **Note**  Immediately after creating a shared workspace site and then creating the active document in the site, the active document is closed and then reopened so that the copy of the active document that the user sees is the one located in the site. If the active document was saved prior to invoking the  **CreateNew** method, that copy of the document is unavailable for the period of time while the new copy is created. This causes an exception for any code that tries to access the saved copy during the creation time period. One workaround is to impose a short delay (suggested 15 seconds or more) before attempting to access the active document from any script. In addition, any cached object that points to the local document msut be updated to point to the document in the shared workspace site.


## Example

The following example creates a shared workspace site at the URL http://server/sites/mysite/, names the workspace "My Shared Budget Document", and adds the active document to the site. The  **URL** property of the new shared workspace site returns http://server/sites/mysite/My%20Shared%20Budget%20Document/, the **Name** property returns "My Shared Budget Document, and **Count** property of the **SharedWorkspaceFiles** collection shows a single file.


```
   Dim sws As Office.SharedWorkspace 
    Dim strSWSInfo As String 
    Set sws = ActiveWorkbook.SharedWorkspace 
    sws.CreateNew "http://server/sites/mysite/", "My Shared Budget Document" 
    strSWSInfo = "Name: " &amp; sws.Name &amp; vbCrLf &amp; _ 
        "URL: " &amp; sws.URL &amp; vbCrLf &amp; _ 
        "File(s): " &amp; sws.Files.Count 
    MsgBox strSWSInfo, vbInformation + vbOKOnly, _ 
        "New Shared Workspace Information" 
    Set sws = Nothing 

```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

