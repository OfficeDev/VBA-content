---
title: Explorer.BeforeItemPaste Event (Outlook)
keywords: vbaol11.chm463
f1_keywords:
- vbaol11.chm463
ms.prod: outlook
api_name:
- Outlook.Explorer.BeforeItemPaste
ms.assetid: a6d43429-5309-4b07-7b0b-68cddd2d7e59
ms.date: 06/08/2017
---


# Explorer.BeforeItemPaste Event (Outlook)

Occurs when an Outlook item is pasted.


## Syntax

 _expression_ . **BeforeItemPaste**( **_ClipboardContent_** , **_Target_** , **_Cancel_** )

 _expression_ A variable that represents an **Explorer** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ClipboardContent_|Required| **Variant**|The content to be pasted.|
| _Target_|Required| **Folder**|The destination of the paste.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the item is not deleted.|

## Remarks

This event can be cancelled after it has started.


## Example

The following Microsoft Visual Basic for Applications (VBA) example prompts the user before pasting the contents of the Clipboard to the specified target. If the user clicks  **Yes**, the current content in the Clipboard is copied to the specified target destination. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `Initialize_handler` routine must be called before the event procedure can be called by Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Sub Initalize_Handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_BeforeItemPaste(ClipboardContent As Variant, ByVal Target As Folder, Cancel As Boolean) 
 
 Dim lngAns As Integer 'users' answer 
 
 'Prompt user about paste 
 
 lngAns = MsgBox("Are you sure you want to paste the contents of the clipboard into the " _ 
 
 &; Target.Name &; "?", vbYesNo) 
 
 If lngAns = vbNo Then 
 
 Cancel = True 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

