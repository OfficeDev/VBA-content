---
title: IBlogExtensibility.GetCategories Method (Office)
keywords: vbaof11.chm328008
f1_keywords:
- vbaof11.chm328008
ms.prod: office
api_name:
- Office.IBlogExtensibility.GetCategories
ms.assetid: f263594c-db27-86bd-8597-35a3148a5ea7
ms.date: 06/08/2017
---


# IBlogExtensibility.GetCategories Method (Office)

This method returns the list of blog categories for an account so Microsoft Word can populate the categories dropdown list.


## Syntax

 _expression_. **GetCategories**( **_Account_**, **_ParentWindow_**, **_Document_**, **_userName_**, **_Password_**, **_Categories()_** )

 _expression_ An expression that returns a **IBlogExtensibility** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. |
| _ParentWindow_|Required|**Long**|Represents the HWND of the host window.|
| _Document_|Required|**Object**|The current document.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents user's password stored in the registry account settings.|
| _Categories()_|Required|**String**|A list of categories supported by the provider.|

## Remarks

Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.


## See also


#### Concepts


[IBlogExtensibility Object](iblogextensibility-object-office.md)
#### Other resources


[IBlogExtensibility Object Members](iblogextensibility-members-office.md)

