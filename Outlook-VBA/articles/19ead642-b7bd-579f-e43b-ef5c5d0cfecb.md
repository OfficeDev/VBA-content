
# MailItem.Display Method (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Displays a new  ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)** object for the item.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Display**( **_Modal_**)

 _expression_A variable that represents a  **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Modal|Optional| **Variant**| **True** to make the window modal. The default value is **False**.|

## Remarks
<a name="sectionSection1"> </a>

The  **Display** method is supported for explorer and inspector windows for the sake of backward compatibility. To activate an explorer or inspector window, use the ** [Activate](d7784df0-b595-6f5a-2195-27ad021db6de.md)** method.

If you attempt to open an "unsafe" file system object (or "freedoc" file) by using the Microsoft Outlook object model, you receive the  **E_FAIL** return code in the C or C++ programming languages. In Outlook 2000 and earlier, you could open an "unsafe" file system object by using the **Display** method.


## Example
<a name="sectionSection2"> </a>

This Visual Basic for Applications example displays the first item in the  **Inbox** folder. This example will return an error if the **Inbox** is empty, because you are trying to display a specific item. If there are no items in the folder, a message box will be displayed to inform the user.


 **Note**  The items in the  **Items** collection object are not guaranteed to be in any particular order.


```
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
<a name="sectionSection2"> </a>


#### Concepts


 [MailItem Object](14197346-05d2-0250-fa4c-4a6b07daf25f.md)
#### Other resources


 [MailItem Object Members](1094d7df-ee80-a4b0-5a21-db2979506e6b.md)
