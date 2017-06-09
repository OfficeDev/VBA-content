---
title: ProtectedViewWindows.Open Method (Word)
keywords: vbawd10.chm82313218
f1_keywords:
- vbawd10.chm82313218
ms.prod: word
api_name:
- Word.ProtectedViewWindows.Open
ms.assetid: 38a11e87-bc8e-4a3b-3b0d-aa51eef941b5
ms.date: 06/08/2017
---


# ProtectedViewWindows.Open Method (Word)

Opens the specified document in a new protected view window.


## Syntax

 _expression_ . **Open**( **_FileName_** , **_AddToRecentFiles_** , **_PasswordDocument_** , **_Visible_** , **_OpenAndRepair_** )

 _expression_ An expression that returns a **ProtectedViewWindows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **Variant**|The name of the document (paths are accepted).|
| _AddToRecentFiles_|Optional| **Variant**| **True** to add the file name to the list of recently used files at the bottom of the File menu.|
| _PasswordDocument_|Optional| **Variant**|The password for opening the document.|
| _Visible_|Optional| **Variant**| **True** if the document is opened in a visible window. The default value is **True** .|
| _OpenAndRepair_|Optional| **Variant**| **True** to repair the document to prevent document corruption.|

### Return Value

ProtectedViewWindow


## Remarks

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code.


## Example

The following code example opens a document in a new protected view window.


```
ProtectedViewWindows.Open FileName:="C:\MyFiles\MyDoc.doc" 

```


## See also


#### Concepts


[ProtectedViewWindows Object](protectedviewwindows-object-word.md)

