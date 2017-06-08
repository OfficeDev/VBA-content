---
title: TaskRequestUpdateItem.SaveAs Method (Outlook)
keywords: vbaol11.chm1954
f1_keywords:
- vbaol11.chm1954
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.SaveAs
ms.assetid: 7d40f1b8-d5df-f301-4350-b783c480fe72
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.SaveAs Method (Outlook)

Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.


## Syntax

 _expression_ . **SaveAs**( **_Path_** , **_Type_** )

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path in which to save the item.|
| _Type_|Optional| **Variant**|The file type to save. Can be one of the following  **OlSaveAsType** constants: **olHTML** , **olMSG** , **olRTF** , **olTemplate** , **olDoc** , ** olTXT** , **olVCal** , **olVCard** , **olICal** , or **olMSGUnicode** .|

## Remarks

Also note that even though  **olDoc** is a valid **OlSaveAsType** constant, messages in HTML format cannot be saved in Document format, and the **olDoc** constant works only if Microsoft Word is set up as the default email editor.


## See also


#### Concepts


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

