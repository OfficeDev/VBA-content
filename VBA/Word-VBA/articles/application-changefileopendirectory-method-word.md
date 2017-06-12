---
title: Application.ChangeFileOpenDirectory Method (Word)
keywords: vbawd10.chm158335333
f1_keywords:
- vbawd10.chm158335333
ms.prod: word
api_name:
- Word.Application.ChangeFileOpenDirectory
ms.assetid: 9f044713-6e97-7219-8083-7d7d2cbb1b0f
ms.date: 06/08/2017
---


# Application.ChangeFileOpenDirectory Method (Word)

Sets the folder in which Word searches for documents.


## Syntax

 _expression_ . **ChangeFileOpenDirectory**( **_Path_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path to the folder in which Word searches for documents.|

## Remarks

The specified folder's contents are listed the next time the  **Open** dialog box ( **File** tab) is displayed. Word searches the specified folder for documents until the user changes the folder in the **Open** dialog box or the current Word session ends. Use the **[DefaultFilePath](options-defaultfilepath-property-word.md)** property to change the default folder for documents in every Word session.


## Example

This example changes the folder in which Word searches for documents, and then opens a file named "Test.doc."


```
ChangeFileOpenDirectory "C:\Documents" 
Documents.Open FileName:="Test.doc"
```

This example changes the folder in which Word searches for documents, and then displays the Open dialog box.




```
ChangeFileOpenDirectory "C:\" 
Dialogs(wdDialogFileOpen).Show
```


## See also


#### Concepts


[Application Object](application-object-word.md)

