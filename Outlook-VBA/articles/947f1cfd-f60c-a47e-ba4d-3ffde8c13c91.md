
# AppointmentItem.Copy Method (Outlook)

 **Last modified:** July 28, 2015

Creates another instance of an object.

## Syntax

 _expression_. **Copy**

 _expression_A variable that represents an  **AppointmentItem** object.


## Example

This Visual Basic for Applications example creates an e-mail message, sets the  **Subject**to "Speeches", uses the  **Copy** method to copy it, then moves the copy into a newly created e-mail folder named "Saved Mail" within the Inbox folder.


```
Sub CopyItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myNewFolder As Outlook.Folder 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myCopiedItem As Outlook.MailItem 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myNewFolder = myFolder.Folders.Add("Saved Mail", olFolderDrafts) 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 myItem.Subject = "Speeches" 
 
 Set myCopiedItem = myItem.Copy 
 
 myCopiedItem.Move myNewFolder 
 
End Sub
```


## See also


#### Concepts


 [AppointmentItem Object](204a409d-654e-27aa-643a-8344c631b82d.md)
#### Other resources


 [AppointmentItem Object Members](c72c459d-6d3c-7a05-aa4a-b1b767ddc0b2.md)
