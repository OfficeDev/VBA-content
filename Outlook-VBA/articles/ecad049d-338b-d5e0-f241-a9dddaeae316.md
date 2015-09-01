
# Conversation.GetAlwaysMoveToFolder Method (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [Folder](3cf6cda8-6d70-666e-2643-9d9c5b9cacfc.md)** object that indicates the folder in the specified delivery store to which new items that arrive in the conversation are always moved.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **GetAlwaysMoveToFolder**( **_Store_**)

 _expression_A variable that represents a  ** [Conversation](2705d38a-ebc0-e5a7-208b-ffe1f5446b1b.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Store|Required| ** [Store](1eb22fe9-8849-7476-5388-2515b48591b9.md)**|The store where the folder to which conversation items are moved resides.|

### Return Value

A  **Folder** object in the specified store to which all new items that arrive in the conversation are always moved.


## Remarks
<a name="sectionSection1"> </a>

If the  _Store_ parameter represents a non-delivery store such as an archive .pst store, the **GetAlwaysMoveToFolder** method returns a **Folder** object that applies to conversation items on the default delivery store.

If no folder, other than the  **Deleted Items** folder, has been specified to always move conversation items into, the **GetAlwaysMoveToFolder** method returns **Null** ( **Nothing** in Visual Basic).


## Example
<a name="sectionSection2"> </a>

The following Microsoft Visual Basic for Application (VBA) example shows how to find the folder into which new items that arrive in the conversation of the first mail item displayed in the Reading Pane are always moved. The code example,  `DemoGetAlwaysMoveToFolder`, verifies that conversations are enabled in the store for the selected mail item, obtains the conversation object for that mail item if a conversation exists, uses  **GetAlwaysMoveToFolder** to obtain the folder, and displays the folder name.


```
Sub DemoGetAlwaysMoveToFolder() 
 
 Dim oMail As Outlook.MailItem 
 
 Dim oConv As Outlook.Conversation 
 
 Dim oStore As Outlook.Store 
 
 
 
 ' Get Item displayed in Reading Pane. 
 
 Set oMail = ActiveExplorer.Selection(1) 
 
 Set oStore = oMail.Parent.Store 
 
 If oStore.IsConversationEnabled Then 
 
 Set oConv = oMail.GetConversation 
 
 If Not (oConv Is Nothing) Then 
 
 Dim oFolder As Outlook.folder 
 
 Set oFolder = _ 
 
 oConv.GetAlwaysMoveToFolder(oStore) 
 
 If Not (oFolder Is Nothing) Then 
 
 Debug.Print "MoveToFolder: " &amp; oFolder.name 
 
 Else 
 
 Debug.Print "MoveToFolder action not set" 
 
 End If 
 
 End If 
 
 End If 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Conversation Object](2705d38a-ebc0-e5a7-208b-ffe1f5446b1b.md)
#### Other resources


 [Conversation Object Members](09ff1e8e-7c5a-0b1e-e8e2-e259f66f71c8.md)
