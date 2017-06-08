---
title: ReportItem.SaveAs Method (Outlook)
keywords: vbaol11.chm1671
f1_keywords:
- vbaol11.chm1671
ms.prod: outlook
api_name:
- Outlook.ReportItem.SaveAs
ms.assetid: 70497e98-0b4d-266b-10c1-c340a14e82c9
ms.date: 06/08/2017
---


# ReportItem.SaveAs Method (Outlook)

Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.


## Syntax

 _expression_ . **SaveAs**( **_Path_** , **_Type_** )

 _expression_ A variable that represents a **ReportItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path in which to save the item.|
| _Type_|Optional| **Variant**|The file type to save. Can be one of the following  **OlSaveAsType** constants: **olHTML** , **olMSG** , **olRTF** , **olTemplate** , **olDoc** , ** olTXT** , **olVCal** , **olVCard** , **olICal** , or **olMSGUnicode** .|

## Remarks

Also note that even though  **olDoc** is a valid **OlSaveAsType** constant, messages in HTML format cannot be saved in Document format, and the **olDoc** constant works only if Microsoft Word is set up as the default email editor.


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

