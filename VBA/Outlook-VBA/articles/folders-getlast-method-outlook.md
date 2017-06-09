---
title: Folders.GetLast Method (Outlook)
keywords: vbaol11.chm48
f1_keywords:
- vbaol11.chm48
ms.prod: outlook
api_name:
- Outlook.Folders.GetLast
ms.assetid: 6d981844-3ac0-c6f9-b2ee-9cf495ab6488
ms.date: 06/08/2017
---


# Folders.GetLast Method (Outlook)

Returns the last object in the  **[Folders](folders-object-outlook.md)** collection.


## Syntax

 _expression_ . **GetLast**

 _expression_ A variable that represents a **Folders** object.


### Return Value

A  **[Folder](folder-object-outlook.md)** object that represents the last object contained by the collection.


## Remarks

It returns  **Nothing** if no last object exists, for example, if the collection is empty.To ensure correct operation of the **[GetFirst](folders-getfirst-method-outlook.md)** , **GetLast** , **[GetNext](folders-getnext-method-outlook.md)** , and **[GetPrevious](folders-getprevious-method-outlook.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious** . To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## Example

The following Visual Basic for Applications example searches the subfolders of  **Inbox** for a folder called **MyPersonalEmails** and displays a message to the user. If you do not have a subfolder called **MyPersonalEmails** in your **Inbox** folder, the example will display nothing.


```vb
Sub TestGetLast() 
 
 Dim nsp As Outlook.NameSpace 
 
 Dim mpf As Outlook.Folder 
 
 Dim mpfSubFolder As Outlook.Folder 
 
 Dim flds As Outlook.Folders 
 
 Dim idx As Integer 
 
 
 
 Set nsp = Application.GetNamespace("MAPI") 
 
 Set mpf = nsp.GetDefaultFolder(olFolderInbox) 
 
 Set flds = mpf.Folders 
 
 Set mpfSubFolder = flds.GetLast 
 
 Do While Not mpfSubFolder Is Nothing 
 
 If mpfSubFolder.Name = "MyPersonalEmails" Then 
 
 MsgBox "The folder was found." 
 
 Exit Do 
 
 End If 
 
 Set mpfSubFolder = flds.GetPrevious 
 
 Loop 
 
End Sub
```


## See also


#### Concepts


[Folders Object](folders-object-outlook.md)

