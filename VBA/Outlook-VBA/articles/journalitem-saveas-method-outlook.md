---
title: JournalItem.SaveAs Method (Outlook)
keywords: vbaol11.chm1264
f1_keywords:
- vbaol11.chm1264
ms.prod: outlook
api_name:
- Outlook.JournalItem.SaveAs
ms.assetid: f23ffcb6-8c08-1dc0-d491-6389303934ec
ms.date: 06/08/2017
---


# JournalItem.SaveAs Method (Outlook)

Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.


## Syntax

 _expression_ . **SaveAs**( **_Path_** , **_Type_** )

 _expression_ A variable that represents a **JournalItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path in which to save the item.|
| _Type_|Optional| **Variant**|The file type to save. Can be one of the following  **OlSaveAsType** constants: **olHTML** , **olMSG** , **olRTF** , **olTemplate** , **olDoc** , ** olTXT** , **olVCal** , **olVCard** , **olICal** , or **olMSGUnicode** .|

## Remarks

Also note that even though  **olDoc** is a valid **OlSaveAsType** constant, messages in HTML format cannot be saved in Document format, and the **olDoc** constant works only if Microsoft Word is set up as the default email editor.


## Example

This Visual Basic for Applications (VBA) example uses the  **SaveAs** method to save the currently open item as a text file in the Documents folder, using the subject as the file name. To run this example, make sure a mail item in plain text format is open in the active window.


```vb
Sub SaveAsTXT() 
 Dim myItem As Outlook.Inspector 
 Dim objItem As Object 
 
 Set myItem = Application.ActiveInspector 
 If Not TypeName(myItem) = "Nothing" Then 
 Set objItem = myItem.CurrentItem 
 strname = objItem.Subject 
 'Prompt the user for confirmation 
 Dim strPrompt As String 
 strPrompt = "Are you sure you want to save the item? " &; _ 
 "If a file with the same name already exists, " &; _ 
 "it will be overwritten with this copy of the file." 
 If MsgBox(strPrompt, vbYesNo + vbQuestion) = vbYes Then 
 objItem.SaveAs Environ("HOMEPATH") &; "\My Documents\" &; strname &; ".txt", olTXT 
 End If 
 Else 
 MsgBox "There is no current active inspector." 
 End If 
End Sub
```

This Visual Basic for Applications example shows you how to create a template using the  **SaveAs** method.




```vb
Sub CreateTemplate() 
 Dim MyItem As Outlook.JournalItem 
 
 Set MyItem = Application.CreateItem(olJournalItem) 
 MyItem.Subject = "Status Report" 
 MyItem.Display 
 MyItem.SaveAs Environ("HOMEPATH") &; "\My Documents\statusrep.oft", OlSaveAsType.olTemplate 
End Sub
```


## See also


#### Concepts


[JournalItem Object](journalitem-object-outlook.md)

