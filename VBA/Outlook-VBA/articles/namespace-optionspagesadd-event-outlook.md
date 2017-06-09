---
title: NameSpace.OptionsPagesAdd Event (Outlook)
keywords: vbaol11.chm319
f1_keywords:
- vbaol11.chm319
ms.prod: outlook
api_name:
- Outlook.NameSpace.OptionsPagesAdd
ms.assetid: 3f4920bd-ab22-90a7-490a-67122dac6c51
ms.date: 06/08/2017
---


# NameSpace.OptionsPagesAdd Event (Outlook)

Occurs whenever the  **Properties** dialog box for a folder is opened.


## Syntax

 _expression_ . **OptionsPagesAdd**( **_Pages_** , **_Folder_** )

 _expression_ A variable that represents a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pages_|Required| **[PropertyPages](propertypages-object-outlook.md)**|The collection of property pages that have been added to the dialog box. This collection includes only custom property pages. It does not include standard Microsoft Outlook property pages.|
| _Folder_|Required| **[Folder](folder-object-outlook.md)**|This argument is only used with the  **Folder** object. The **Folder** object for which the **Properties** dialog box is being opened.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).

You can open the  **Properties** dialog box for a folder by right-clicking the folder and selecting **Properties**. 

Your program handles this event to add a custom property page. The property page will be added to  **Properties** dialog box of the specified folder. When the event fires, the **PropertyPages** collection object identified by _Pages_ contains the property pages that have been added prior to the event handler being called. To add your property page to the collection, use the **[Add](propertypages-add-method-outlook.md)** method of the **PropertyPages** collection before exiting the event handler.


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

