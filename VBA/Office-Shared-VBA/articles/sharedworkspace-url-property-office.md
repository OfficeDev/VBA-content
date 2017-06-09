---
title: SharedWorkspace.URL Property (Office)
keywords: vbaof11.chm276011
f1_keywords:
- vbaof11.chm276011
ms.prod: office
api_name:
- Office.SharedWorkspace.URL
ms.assetid: e60e6706-d3f3-1a47-2b8a-82c5d52ddac5
ms.date: 06/08/2017
---


# SharedWorkspace.URL Property (Office)

Gets the top-level Uniform Resource Locator (URL) of the shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **URL**

 _expression_ A variable that represents a **SharedWorkspace** object.


### Return Value

String


## Remarks

The URL property returns the address of the shared workspace in this format:  `http://server/sites/user/workspace/`. The URL property returns a URL-encoded string. For example, a space in the folder name is represented by %20. Use a simple function like the following example to replace this escaped character with a space. `Private Function URLDecode(URLtoDecode As String) As String URLDecode = Replace(URLtoDecode, "%20", " ") End Function`


## Example

The following example displays the base URL of the shared workspace.


```
 MsgBox "URL: " &amp; ActiveWorkbook.SharedWorkspaceLink.URL, _ 
        vbInformation + vbOKOnly, "Shared Workspace URL" 

```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

