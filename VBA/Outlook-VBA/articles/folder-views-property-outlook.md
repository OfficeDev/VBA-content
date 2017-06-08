---
title: Folder.Views Property (Outlook)
keywords: vbaol11.chm2011
f1_keywords:
- vbaol11.chm2011
ms.prod: outlook
api_name:
- Outlook.Folder.Views
ms.assetid: 24ef613a-9832-032c-4e68-1001a0385b11
ms.date: 06/08/2017
---


# Folder.Views Property (Outlook)

Returns the  **[Views](views-object-outlook.md)** collection object of the **[Folder](folder-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **Views**

 _expression_ A variable that represents a **Folder** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates an instance of the  **Views** collection and displays the XML definition of a view called "Table View". If the view does not exist, it creates one.


```vb
Sub DisplayViewDef() 
 
 'Displays the XML definition of a View object 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objViews As Outlook.Views 
 
 Dim objView As Outlook.View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 'Return a view called Table View if it already exists, else create one 
 
 Set objView = objViews.Item("Table View") 
 
 If objView Is Nothing Then 
 
 Set objView = objViews.Add("Table View", olTableView, _ 
 
 olViewSaveOptionAllFoldersOfType) 
 
 End If 
 
 MsgBox objView.XML 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

