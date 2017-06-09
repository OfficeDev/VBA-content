---
title: IBlogExtensibility.RepublishPost Method (Office)
keywords: vbaof11.chm328007
f1_keywords:
- vbaof11.chm328007
ms.prod: office
api_name:
- Office.IBlogExtensibility.RepublishPost
ms.assetid: 1e701746-f63b-68a3-6a5c-75b78942d380
ms.date: 06/08/2017
---


# IBlogExtensibility.RepublishPost Method (Office)

Hands off the current post so it can be republished by the provider.


## Syntax

 _expression_. **RepublishPost**( **_Account_**, **_ParentWindow_**, **_Document_**, **_userName_**, **_Password_**, **_PostID_**, **_xHTML_**, **_Title_**, **_DateTime_**, **_Categories()_**, **_Draft_**, **_PublishMessage_** )

 _expression_ An expression that returns a **IBlogExtensibility** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents user's password stored in the registry account settings.|
| _PostID_|Required|**String**|The ID of the original post.|
| _xHTML_|Required|**String**|Represents the xHTML of the current document.|
| _Title_|Required|**String**|The title of the post.|
| _DateTime_|Required|**String**|The date the entry was posted.|
| _Categories()_|Required|**String**|A list of categories supported by the provider.|
| _Draft_|Required|**Boolean**|Specifies whether this is a draft version of the post.|
| _PublishMessage_|Required|**String**|Specifies what is displayed in the publish bar.|

## See also


#### Concepts


[IBlogExtensibility Object](iblogextensibility-object-office.md)
#### Other resources


[IBlogExtensibility Object Members](iblogextensibility-members-office.md)

