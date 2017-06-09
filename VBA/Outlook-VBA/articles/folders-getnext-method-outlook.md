---
title: Folders.GetNext Method (Outlook)
keywords: vbaol11.chm49
f1_keywords:
- vbaol11.chm49
ms.prod: outlook
api_name:
- Outlook.Folders.GetNext
ms.assetid: 5c2de8b2-b251-1983-a10b-1945abc38709
ms.date: 06/08/2017
---


# Folders.GetNext Method (Outlook)

Returns the next object in the  **[Folders](folders-object-outlook.md)** collection.


## Syntax

 _expression_ . **GetNext**

 _expression_ A variable that represents a **Folders** object.


### Return Value

A  **[Folder](folder-object-outlook.md)** object that represents the next object contained by the collection.


## Remarks

It returns  **Nothing** if no next object exists, for example, if already positioned at the end of the collection.To ensure correct operation of the **[GetFirst](folders-getfirst-method-outlook.md)** , **[GetLast](folders-getlast-method-outlook.md)** , **GetNext** , and **[GetPrevious](folders-getprevious-method-outlook.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious** . To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## Example

The following Visual Basic for Applications example searches the subfolders of  **Inbox** for a folder called **MyPersonalEmails** and displays a message to the user. If you do not have a subfolder called **MyPersonalEmails** in your **Inbox** folder, the example will display nothing.


```vb
Sub TestGetNext() 
 
 Dim nsp As Outlook.NameSpace 
 
 Dim mpf As Outlook.Folder 
 
 Dim mpfSubFolder As Outlook.Folder 
 
 Dim flds As Outlook.Folders 
 
 Dim idx As Integer 
 
 
 
 Set nsp = Application.GetNamespace("MAPI") 
 
 Set mpf = nsp.GetDefaultFolder(olFolderInbox) 
 
 Set flds = mpf.Folders 
 
 Set mpfSubFolder = flds.GetFirst 
 
 Do While Not mpfSubFolder Is Nothing 
 
 If mpfSubFolder.Name = "MyPersonalEmails" Then 
 
 MsgBox "The folder was found." 
 
 Exit Do 
 
 End If 
 
 Set mpfSubFolder = flds.GetNext 
 
 Loop 
 
End Sub
```


## See also


#### Concepts


[Folders Object](folders-object-outlook.md)

