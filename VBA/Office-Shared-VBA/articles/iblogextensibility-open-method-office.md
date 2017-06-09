---
title: IBlogExtensibility.Open Method (Office)
keywords: vbaof11.chm328005
f1_keywords:
- vbaof11.chm328005
ms.prod: office
api_name:
- Office.IBlogExtensibility.Open
ms.assetid: 34bae5c9-cc29-b1b8-746b-bc2630cf8bc0
ms.date: 06/08/2017
---


# IBlogExtensibility.Open Method (Office)

Opens the blog specified by the blog ID. It is called by the  **Open Existing Post** dialog based on the item selected by the user.


## Syntax

 _expression_. **Open**( **_Account_**, **_PostID_**, **_ParentWindow_**, **_userName_**, **_Password_**, **_xHTML_**, **_Title_**, **_DatePosted_**, **_Categories()_** )

 _expression_ An expression that returns a **IBlogExtensibility** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _PostID_|Required|**String**|The ID of the post.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window Microsoft Word is calling from.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents user's password stored in the registry account settings.|
| _xHTML_|Required|**String**|Represents the xHTML of the current document.|
| _Title_|Required|**String**|The title of the post.|
| _DatePosted_|Required|**String**|The date the entry was posted.|
| _Categories()_|Required|**String**|A list of categories supported by the provider.|

## See also


#### Concepts


[IBlogExtensibility Object](iblogextensibility-object-office.md)
#### Other resources


[IBlogExtensibility Object Members](iblogextensibility-members-office.md)

