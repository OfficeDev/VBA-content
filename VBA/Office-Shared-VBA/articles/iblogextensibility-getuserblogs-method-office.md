---
title: IBlogExtensibility.GetUserBlogs Method (Office)
keywords: vbaof11.chm328003
f1_keywords:
- vbaof11.chm328003
ms.prod: office
api_name:
- Office.IBlogExtensibility.GetUserBlogs
ms.assetid: 00e76f3d-59f2-8580-6f7e-6df8fe51d345
ms.date: 06/08/2017
---


# IBlogExtensibility.GetUserBlogs Method (Office)

Returns the list and details of user blogs associated with the specified account.


## Syntax

 _expression_. **GetUserBlogs**( **_Account_**, **_ParentWindow_**, **_Document_**, **_userName_**, **_Password_**, **_BlogNames()_**, **_BlogIDs()_**, **_BlogURLs()_** )

 _expression_ An expression that returns a **IBlogExtensibility** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents user's password stored in the registry account settings.|
| _BlogNames()_|Required|**String**|Contains all blog names under the current account.|
| _BlogIDs()_|Required|**String**|Contains all blog IDs under the current account.|
| _BlogURLs()_|Required|**String**|Contains all blog URLs under the current account.|

## See also


#### Concepts


[IBlogExtensibility Object](iblogextensibility-object-office.md)
#### Other resources


[IBlogExtensibility Object Members](iblogextensibility-members-office.md)

