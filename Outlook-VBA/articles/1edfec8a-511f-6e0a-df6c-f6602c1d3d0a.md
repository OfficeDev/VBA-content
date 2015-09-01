
# RemoteItem.MarkForDownload Property (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets an  ** [OlRemoteStatus](2df0404c-26c9-87d4-6916-d75aff8e3fbc.md)**constant that determines the status of an item once it is received by a remote user. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **MarkForDownload**

 _expression_A variable that represents a  **RemoteItem** object.


## Remarks
<a name="sectionSection1"> </a>

This property gives remote users with less-than-ideal data-transfer capabilities increased messaging flexibility.


## Example
<a name="sectionSection2"> </a>

The following example searches through the user's  **Inbox** for items that have not yet been fully downloaded. If any items are found that are not fully downloaded, a message is displayed and the item is marked for download.


```
Sub DownloadItems() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim obj As Object 
 
 Dim i As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 'Loop all items in the Inbox folder 
 
 For i = 1 To mpfInbox.Items.Count 
 
 Set obj = mpfInbox.Items.Item(i) 
 
 'Verify if the state of the item is olHeaderOnly 
 
 If obj.DownloadState = olHeaderOnly Then 
 
 MsgBox ("This item has not been fully downloaded.") 
 
 'Mark the item to be downloaded. 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 End If 
 
 Next 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [RemoteItem Object](6302aaff-cdcf-4d86-60f1-4bed15540d9f.md)
#### Other resources


 [RemoteItem Object Members](15c0872e-88cc-9b9b-c31e-c15d6971e6e0.md)
