---
title: Folder.Display Method (Outlook)
keywords: vbaol11.chm1996
f1_keywords:
- vbaol11.chm1996
ms.prod: outlook
api_name:
- Outlook.Folder.Display
ms.assetid: cde389e0-5ec9-8261-5ec0-9a5ba4f8776d
ms.date: 06/08/2017
---


# Folder.Display Method (Outlook)

Displays a new  **[Explorer](explorer-object-outlook.md)** object for the folder.


## Syntax

 _expression_ . **Display**()

 _expression_ A variable that represents a **Folder** object.


## Example

This Visual Basic for Applications (VBA) example uses the  **Display** method to display the default Inbox folder. This example will not return an error, even if there are no items in the Inbox, because you are not asking for the display of a specific item.


```vb
Sub DisplayInbox() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 myFolder.Display 
 
End Sub
```

This Visual Basic for Applications example displays the first item in the Inbox folder. This example will return an error if the Inbox is empty, because you are trying to display a specific item. If there are no items in the folder, a message box will be displayed to inform the user.


 **Note**  The items in the  **Items** collection object are not guaranteed to be in any particular order.




```vb
Sub DisplayFirstItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 On Error GoTo ErrorHandler 
 
 myFolder.Items(1).Display 
 
 Exit Sub 
 
ErrorHandler: 
 
 MsgBox "There are no items to display." 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

