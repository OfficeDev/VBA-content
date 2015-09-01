
# Store.GetSearchFolders Method (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [Folders](0c814c3c-74fc-414c-982d-a0097fcb35c2.md)** collection object that represents the search folders defined for the ** [Store](1eb22fe9-8849-7476-5388-2515b48591b9.md)** object.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **GetSearchFolders**

 _expression_A variable that represents a  **Store** object.


### Return Value

A  **Folders** collection object that represents all the search folders for the **Store** object.


## Remarks
<a name="sectionSection1"> </a>

 **GetSearchFolders** returns all the visible active search folders for the **Store**. It does not return uninitialized or aged out search folders.

 **GetSearchFolders** returns a **Folders** collection object with ** [Folders.Count](b1884cc1-5b50-0ea8-315a-3616d11db0e6.md)** equal zero (0) if no search folders have been defined for the **Store**.

For a  **Folders** collection object that represents a collection of search folders, ** [Folders.Parent](4fe483ec-7e6e-ca82-8a1d-d039a7b9e89c.md)** returns the same object as ** [Store.GetRootFolder](09da4d57-c33d-6946-cc21-7233e89efb10.md)**.  ** [Folder.Folders](41464c32-023e-9079-4f24-51586305325c.md)** returns **Null** ( **Nothing** in Visual Basic).


## Example
<a name="sectionSection2"> </a>

The following code sample in Microsoft Visual Basic for Applications (VBA) enumerates the search folders on all stores for the current session.


```
Sub EnumerateSearchFoldersInStores() 
 
 Dim colStores As Outlook.Stores 
 
 Dim oStore As Outlook.Store 
 
 Dim oSearchFolders As Outlook.folders 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 On Error Resume Next 
 
 Set colStores = Application.Session.Stores 
 
 For Each oStore In colStores 
 
 Set oSearchFolders = oStore.GetSearchFolders 
 
 For Each oFolder In oSearchFolders 
 
 Debug.Print (oFolder.FolderPath) 
 
 Next 
 
 Next 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Store Object](1eb22fe9-8849-7476-5388-2515b48591b9.md)
#### Other resources


 [Store Object Members](84c1d423-e507-0b3b-6570-33829b94be04.md)
