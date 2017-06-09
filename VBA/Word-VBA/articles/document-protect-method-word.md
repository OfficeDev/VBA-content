---
title: Document.Protect Method (Word)
keywords: vbawd10.chm158007763
f1_keywords:
- vbawd10.chm158007763
ms.prod: word
ms.assetid: 727bafe9-48ea-6b2f-2262-778f66487cbd
ms.date: 06/08/2017
---


# Document.Protect Method (Word)

Protects the specified document from unauthorized changes.


## Syntax

 _expression_ . **Protect**_(Type,_ _NoReset,_ _Password,_ _UseIRM,_ _EnforceStyleLock)_

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _Type_|Required| **WdProtectionType**|The type of protection to apply.|
| _NoReset_|Optional|VARIANT| **False** to reset form fields to their default values; **True** to retain the current form field values if the document is protected. If _Type_ is not **wdAllowOnlyFormFields**,  _NoReset_ is ignored.|
| _Password_|Optional|VARIANT|If supplied, the password to be able to edit the document, or to change or remove protection.|
| _UseIRM_|Optional|VARIANT|Specifies whether to use Information Rights Management (IRM) when protecting the document from changes.|
| _EnforceStyleLock_|Optional|VARIANT|Specifies whether formatting restrictions are enforced for a protected document.|
| _Type_|Required|WDPROTECTIONTYPE||
| _NoReset_|Optional|VARIANT||
| _Password_|Optional|VARIANT||
| _UseIRM_|Optional|VARIANT||
| _EnforceStyleLock_|Optional|VARIANT||

### Return value

 **VOID**


## See also


#### Concepts


[Document Object](document-object-word.md)

