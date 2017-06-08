---
title: Working with forms in the Visual Basic Editor
keywords: vbaol11.chm5274251
f1_keywords:
- vbaol11.chm5274251
ms.prod: outlook
ms.assetid: b98ed8f2-32ae-9868-ea65-5e6fa7cc34f2
ms.date: 06/08/2017
---


# Working with forms in the Visual Basic Editor

You can use the Visual Basic Editor to design a form that allows your users to interact with your Microsoft Visual Basic for Applications (VBA) program. Unlike an Outlook form, a Visual Basic for Applications form is not used to display an Outlook item, nor can a control on a Visual Basic for Applications form be bound to an item field.

Your Visual Basic for Applications program can use a Visual Basic for Applications user form to gather information from your users; your program can then use this information to set properties of new or existing Outlook items. For example, a program that creates a boilerplate mail message could use a Visual Basic for Applications form to allow the user to enter the specific information for the message to be sent. When the user closes the form, the program uses the information in the form to set the properties of the mail message and then sends the message.

The following sample uses the text in two text boxes to add information to a message before sending it.




```vb
Private Sub CommandButton1_Click() 
 Dim myMail As Outlook.MailItem 
 Set myMail = Application.CreateItem(olMailItem) 
 With myMail 
 .To = TextBox1.Text 
 .Subject = "Book overdue: " &; TextBox2.Text 
 .Body = "Please return this book as soon as possible." 
 End With 
 myMail.Send 
End Sub
```

You can also use controls to display information about Outlook items, folders, and other features of the Outlook object model. The following example shows how to fill a combo box control with the subjects of the items in the user's Inbox.



```vb
Dim myItems As Outlook.Items 
Set myItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items 
For x = 1 To myItems.Count 
 ComboBox1.AddItem myItems.Item(x).Subject 
Next x
```

For more information about creating and using forms in the Visual Basic Editor, see the Visual Basic Editor Help.

