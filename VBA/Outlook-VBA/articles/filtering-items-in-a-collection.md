---
title: Filtering Items in a Collection
keywords: olfm10.chm3077113
f1_keywords:
- olfm10.chm3077113
ms.prod: outlook
ms.assetid: f46c42a1-73c9-ecee-934a-53db0a1f3368
ms.date: 06/08/2017
---


# Filtering Items in a Collection

You can use the Microsoft Outlook object model to return information about all items in a folder. Often, however, the desired objective is to search for a specific item or to retrieve a subset of the items in the folder. Consider the following examples:


- You are developing a Microsoft Access database. When the user enters a new contact record, you want to give the user the ability to click a button to check whether a contact with the same name already exists in Outlook. If a match is found, you can retrieve all the fields for the contact and automatically fill in the Access database record. In this situation, if the user filled in the first and last name fields on the Access form, you can use the  ** [Items.Find](items-find-method-outlook.md)** method in the Outlook object model to search for a match against the Outlook Full Name field. If you want to make sure there are no additional contacts in Outlook with the same name, you can then use the ** [Items.FindNext](items-findnext-method-outlook.md)** method to conduct the same search again. Note that if you do not require the search results to contain values for all the built-in properties of the item, you should use the **[FindRow](table-findrow-method-outlook.md)** and **[Restrict](table-restrict-method-outlook.md)** methods of the **[Table](table-object-outlook.md)** object for better search performance. For more information on searching and filtering items using the Outlook object model, see [Enumerating, Searching, and Filtering Items in a Folder](enumerating-searching-and-filtering-items-in-a-folder.md).
    
- You are writing a Microsoft Visual Basic program to automatically schedule appointments in users' calendars. In order to do this, you need to retrieve a user's appointments for a given day. In this case, you would use the  ** [Items.Restrict](items-restrict-method-outlook.md)** or **Table.Restrict** method to retrieve all appointments that fall on a particular day.
    

While the  **Items.Find**,  **Items.Restrict**,  **Table.FindRow**, and  **Table.Restrict** methods perform similar search and filter functions, **Items.Find** supports only the Microsoft Jet syntax whereas the others support both the Jet syntax and the DAV Searching and Locating (DASL) syntax. For more information on these syntaxes, see [Filtering Items](filtering-items.md).


