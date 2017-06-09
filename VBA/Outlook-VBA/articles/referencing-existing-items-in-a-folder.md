---
title: Referencing Existing Items in a Folder
keywords: olfm10.chm3077116
f1_keywords:
- olfm10.chm3077116
ms.prod: outlook
ms.assetid: 8995fcd8-bd03-7987-fa4d-88b2cf321eca
ms.date: 06/08/2017
---


# Referencing Existing Items in a Folder

There are a number of ways you can reference existing items in a folder using Microsoft Visual Basic. This topic provides information about:


- Using a  `For … Next` or `For Each … Next` loop
    
- Using the  **[Items](items-object-outlook.md)** collection
    
- Using the  **[Find](items-find-method-outlook.md)** method
    
- Using the  **[Restrict](items-restrict-method-outlook.md)** method
    

## Using a For…Next or For Each...Next Loop

Typically these statements are used to loop through all of the items in a folder. The  **Items** collection contains all the items in a particular folder, and you can specify which item to reference by using an index with the **Items** collection. This is typically used with the `For i = 1 to n` programming construct.

You can use  `For Each...Next` to loop through the items in the collection without specifying an index. Both approaches achieve the same result.

The following examples use  `For…Next` to loop through all the contacts in the Contacts folder and display the Full Name field in a dialog box.




```vb
' Microsoft Visual Basic for Applications code example. 
Set olns = Application.GetNameSpace("MAPI") 
' Set MyFolder to the default contacts folder. 
Set MyFolder = olns.GetDefaultFolder(olFolderContacts) 
' Get the number of items in the folder. 
NumItems = MyFolder.Items.Count 
' Set MyItem to the collection of items in the folder. 
Set myItems = myFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'") 
' Loop through all of the items in the folder. 
For I = 1 to NumItems 
   MsgBox MyItems(I).FullName 
Next 

```




```vb
' Visual Basic Scripting Edition code example. 
Set olns = Item.Application.GetNameSpace("MAPI") 
' Set MyFolder to the default contacts folder. 
Set MyFolder = olns.GetDefaultFolder(10) 
' Get the number of items in the folder. 
NumItems = MyFolder.Items.Count 
' Set MyItem to the collection of items in the folder. 
Set myItems = myFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'") 
' Loop through all of the items in the folder. 
For I = 1 to NumItems 
   MsgBox MyItems(I).FullName 
Next
```

The following examples use  `For Each...Next` to achieve the same result as the preceding examples:




```vb
' Visual Basic/Visual Basic for Applications code example. 
Set olns = Application.GetNameSpace("MAPI") 
' Set MyFolder to the default contacts folder. 
Set MyFolder = olns.GetDefaultFolder(olFolderContacts) 
' Set MyItems to the collection of items in the folder. 
Set myItems = myFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'") 
For Each SpecificItem in MyItems 
   MsgBox SpecificItem.FullName 
Next
```




```vb
' VBScript code example. 
Set olns = Item.Application.GetNameSpace("MAPI") 
' Set MyFolder to the default contacts folder. 
Set MyFolder = olns.GetDefaultFolder(10) 
' Set MyItem to the collection of items in the folder. 
Set myItems = myFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'") 
For Each SpecificItem in MyItems 
   MsgBox SpecificItem.FullName 
Next
```


## Using the Items Collection

You can also use the  **Items** collection and specify a text string that matches the Subject field of an item. The following examples display an item in the Inbox whose subject contains "Please help on Friday!"


```vb
' Visual Basic/Visual Basic for Applications code example. 
Set olns = Application.GetNameSpace("MAPI") 
' Set MyFolder to the default Inbox. 
Set MyFolder = olns.GetDefaultFolder(olFolderInbox) 
Set MyItem = MyFolder.Items("Please help on Friday!") 
MyItem.Display 

```


```vb
' VBScript code example. 
Set olns = Item.Application.GetNameSpace("MAPI") 
' Set MyFolder to the default Inbox. 
Set MyFolder = olns.GetDefaultFolder(6) 
Set MyItem = MyFolder.Items("Please help on Friday!") 
MyItem.Display
```


## Using the Find Method

Use the  **Find** method to search for an item in a folder based on the value of one of its fields. If the search is successful, you can then use the **[FindNext](items-findnext-method-outlook.md)** method to check for additional items that meet the same search criteria.

The following examples search to see if you have any high priority tasks.




```vb
' Visual Basic/Visual Basic for Applications code example. 
Set olns = Application.GetNamespace("MAPI") 
Set myFolder = olns.GetDefaultFolder(olFolderTasks) 
Set MyTasks = myFolder.Items 
' Importance corresponds to Priority on the task form. 
Set MyTask = MyTasks.Find("[Importance] = ""High""") 
If MyTask Is Nothing Then ' the Find failed 
   MsgBox "Nothing important. Go party!" 
Else 
   MsgBox "You have something important to do!" 
End If
```




```vb
' VBScript code example. 
Set olns = Item.Application.GetNamespace("MAPI") 
Set myFolder = olns.GetDefaultFolder(13) 
Set MyTasks = myFolder.Items 
' Importance corresponds to Priority on the task form. 
Set MyTask = MyTasks.Find("[Importance] = ""High""") 
If MyTask Is Nothing Then ' the Find failed 
   MsgBox "Nothing important. Go party!" 
Else 
   MsgBox "You have something important to do!" 
End If
```


## Using the Restrict Method

The  **Restrict** method is similar to the **Find** method, but instead of returning a single item, it returns a collection of items that meet the search criteria. For example, you could use this method to find all contacts that work at the same company.

The following examples display all of the contacts that work at ProseWare Corporation:




```vb
' Automation code example. 
Set olns = Application.GetNameSpace("MAPI") 
Set MyFolder = olns.GetDefaultFolder(olFolderContacts) 
Set myItems = myFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'") 
MyClause = "[CompanyName] = ""ProseWare""" 
Set MyPWItems = MyItems.Restrict(MyClause) 
For Each MyItem in MyPWItems 
   MyItem.Display 
Next
```




```vb
' VBScript code example. 
Set olns = Item.Application.GetNameSpace("MAPI") 
Set MyFolder = olns.GetDefaultFolder(10) 
Set myItems = myFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'") 
MyClause = "[CompanyName] = ""ProseWare""" 
Set MyPWItems = MyItems.Restrict(MyClause) 
For Each MyItem in MyPWItems 
   MyItem.Display 
Next
```


