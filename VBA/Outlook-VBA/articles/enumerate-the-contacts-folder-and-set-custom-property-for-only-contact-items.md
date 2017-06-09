---
title: Enumerate the Contacts Folder and Set Custom Property for only Contact Items
ms.prod: outlook
ms.assetid: 6a4cd2e4-a5ec-e55c-0d47-ff618c186c8e
ms.date: 06/08/2017
---


# Enumerate the Contacts Folder and Set Custom Property for only Contact Items

An Outlook folder can contain items of more than one message class. For example, by default, you can create contact items and distribution list items in the Contacts folder. If you want to systematically apply an action to only the contact items or to only the distribution list items in the folder, you must check for the message class for each item in the folder before applying the action.

This topic shows a code sample that uses the message class of an item to identify contact items and sets a user-defined  **Affiliation** field for all contact items in the Contacts folder. The following describes the process:


1. The code sample gets all the items in the default Contacts folder.
    
2. It uses  **[Items.Restrict](items-restrict-method-outlook.md)** to filter contact items from all the items in the default Contacts folder.
    
3. For each contact item, it uses  **[UserProperties.Add](userproperties-add-method-outlook.md)** to add a user-defined field **Affiliation** and sets it based on the existence of a home telephone number. If a home telephone number does not exist for the item, the **Affiliation** property is set to **Business**; otherwise, the field is set to  **Personal**.
    


## Remarks

To run this code sample, place the code in the built-in  **ThisOutlookSession** module. Run the `SetAffiliationForContacts` procedure.

Note that if a field of the name  **Affiliation** already exists, running this example will overwrite it.




```vb
Sub SetAffiliationForContacts() 
 Dim ns As NameSpace 
 Dim foldContact As Folder 
 Dim itemContact As ContactItem 
 Dim colItems As Outlook.Items 
 Dim myProperty As Outlook.UserProperty 
 
 Set ns = Application.GetNamespace("MAPI") 
 Set foldContact = ns.GetDefaultFolder(olFolderContacts) 
 Set colItems = foldContact.Items.Restrict("[MessageClass]='IPM.Contact'") 
 
 For Each itemContact In colItems 
 ' Add user property to contact items 
 Set myProperty = itemContact.UserProperties.Add("Affiliation", olText) 
 If itemContact.HomeTelephoneNumber = "" Then 
 myProperty = "Business" 
 Else 
 myProperty = "Personal" 
 End If 
 itemContact.Save 
 Next 
End Sub
```


