---
title: IBlogPictureExtensibility.CreatePictureAccount Method (Office)
keywords: vbaof11.chm329002
f1_keywords:
- vbaof11.chm329002
ms.prod: office
api_name:
- Office.IBlogPictureExtensibility.CreatePictureAccount
ms.assetid: 8012b234-b8c1-cfc7-7413-b43300fdab76
ms.date: 06/08/2017
---


# IBlogPictureExtensibility.CreatePictureAccount Method (Office)

Allows a picture provider to display the user interface needed to guide the user through setting up a picture account.


## Syntax

 _expression_. **CreatePictureAccount**( **_Account_**, **_BlogProvider_**, **_ParentWindow_**, **_Document_**, **_userName_**, **_Password_** )

 _expression_ An expression that returns a **IBlogPictureExtensibility** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _BlogProvider_|Required|**String**|The ID of the provider.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents user's password stored in the registry account settings.|

## See also


#### Concepts


[IBlogPictureExtensibility Object](iblogpictureextensibility-object-office.md)
#### Other resources


[IBlogPictureExtensibility Object Members](iblogpictureextensibility-members-office.md)

