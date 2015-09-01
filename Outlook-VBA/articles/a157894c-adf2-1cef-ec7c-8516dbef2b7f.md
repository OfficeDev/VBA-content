
# MailItem.SenderEmailAddress Property (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  **String** that represents the e-mail address of the sender of the Outlook item. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SenderEmailAddress**

 _expression_A variable that represents a  **MailItem** object.


## Remarks
<a name="sectionSection1"> </a>

This property corresponds to the MAPI property  **PidTagSenderEmailAddress**.


## Example
<a name="sectionSection2"> </a>

The following Microsoft Visual Basic for Applications (VBA) example loops all items in a folder named Test in the  **Inbox** and sets the yellow flag on items sent by 'someone@example.com'. To run this example without errors, make sure the Test folder exists in the default **Inbox** folder and replace 'someone@example.com' with a valid sender e-mail address in the Test folder.


```
Sub SetFlagIcon() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim obj As Outlook.MailItem 
 
 Dim i As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("Test") 
 
 ' Loop all items in the Inbox\Test Folder 
 
 For i = 1 To mpfInbox.Items.Count 
 
 If mpfInbox.Items(i).Class = olMail Then 
 
 Set obj = mpfInbox.Items.Item(i) 
 
 If obj.SenderEmailAddress = "someone@example.com" Then 
 
 'Set the yellow flag icon 
 
 obj.FlagIcon = olYellowFlagIcon 
 
 obj.Save 
 
 End If 
 
 End If 
 
 Next 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [MailItem Object](14197346-05d2-0250-fa4c-4a6b07daf25f.md)
#### Other resources


 [MailItem Object Members](1094d7df-ee80-a4b0-5a21-db2979506e6b.md)
