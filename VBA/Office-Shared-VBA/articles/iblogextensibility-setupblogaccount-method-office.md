---
title: IBlogExtensibility.SetupBlogAccount Method (Office)
keywords: vbaof11.chm328002
f1_keywords:
- vbaof11.chm328002
ms.prod: office
api_name:
- Office.IBlogExtensibility.SetupBlogAccount
ms.assetid: 98082a55-3e67-7181-2c7d-2c6979c89ab2
ms.date: 06/08/2017
---


# IBlogExtensibility.SetupBlogAccount Method (Office)

Called from the  **Choose Account** dialog when the provider's name is chosen in the **Blog Host** dropdown or when the user requests to change a provider's account in the **Blog Accounts** dialog box.


## Syntax

 _expression_. **SetupBlogAccount**( **_Account_**, **_ParentWindow_**, **_Document_**, **_NewAccount_**, **_ShowPictureUI_** )

 _expression_ An expression that returns a **IBlogExtensibility** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _NewAccount_|Required|**Boolean**|Indicates whether this is a new account.|
| _ShowPictureUI_|Required|**Boolean**|Indicates whether Microsoft Word's picture user interface needs to be displayed.|

## See also


#### Concepts


[IBlogExtensibility Object](iblogextensibility-object-office.md)
#### Other resources


[IBlogExtensibility Object Members](iblogextensibility-members-office.md)

