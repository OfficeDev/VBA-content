---
title: Global.ChangeFileOpenDirectory Method (Word)
keywords: vbawd10.chm163119459
f1_keywords:
- vbawd10.chm163119459
ms.prod: word
api_name:
- Word.Global.ChangeFileOpenDirectory
ms.assetid: 16743466-a8d2-6c4b-063a-eeb8cfb1a2c9
ms.date: 06/08/2017
---


# Global.ChangeFileOpenDirectory Method (Word)

Sets the folder in which Word searches for documents. .


## Syntax

 _expression_ . **ChangeFileOpenDirectory**( **_Path_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path to the folder in which Word searches for documents.|

## Remarks

The contents of the specified folder are listed the next time the  **Open** dialog box ( **File** menu) is displayed.Word searches the specified folder for documents until the user changes the folder in the **Open** dialog box or the current Word session ends. Use the **DefaultFilePath** property to change the default folder for documents in every Word session.


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


[Global Object](global-object-word.md)

