---
title: Explorer.Display Method (Outlook)
keywords: vbaol11.chm2764
f1_keywords:
- vbaol11.chm2764
ms.prod: outlook
api_name:
- Outlook.Explorer.Display
ms.assetid: 3d93be5a-90af-af60-c16a-ec15d87f4d97
ms.date: 06/08/2017
---


# Explorer.Display Method (Outlook)

Displays a new  **[Explorer](explorer-object-outlook.md)** object for the folder.


## Syntax

 _expression_ . **Display**()

 _expression_ A variable that represents an **Explorer** object.


## Remarks

The  **Display** method is supported for explorer and inspector windows for the sake of backward compatibility. To activate an explorer or inspector window, use the **[Activate](inspector-activate-method-outlook.md)** method.

If you attempt to open an "unsafe" file system object (or "freedoc" file) by using the Microsoft Outlook object model, you receive the  **E_FAIL** return code in the C or C++ programming languages. In Outlook 2000 and earlier, you could open an "unsafe" file system object by using the **Display** method.


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

