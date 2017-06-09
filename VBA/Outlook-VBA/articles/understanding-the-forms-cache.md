---
title: Understanding the Forms Cache
ms.prod: outlook
ms.assetid: 0d3baca5-4808-ab86-1ba4-2c676afff5d6
ms.date: 06/08/2017
---


# Understanding the Forms Cache

The forms cache is a folder that is located on a computer's hard disk, and stores a local copy of a standard Microsoft Outlook form. The forms cache does not cache form regions because these form components are already stored on the computer's hard disk.

The forms cache improves the load time of a form because commonly used forms are loaded from the hard disk instead of downloaded from the server. When a form is activated for the first time, the form definition file is copied from its forms library to the Forms folder. The forms cache keeps a temporary copy of the form definition in a subfolder. This subfolder's name roughly matches the name of the form.

The form table, Frmcache.dat, also located in the Forms folder, is used to locate a form and to prevent multiple instances of the same form from being loaded in the cache. When a form is activated, Outlook checks to see if a form with the same message class is already in the cache. If not, it copies the form definition to the cache. In addition, if a change has been made to a form, Outlook copies the new form definition to the cache.

Since Microsoft Office Outlook 2007, Outlook looks for forms in the following order. When it finds a match, Outlook opens the form and does not search further.

1. Forms cached in memory. If you have another item open that uses the same form, Outlook already has that form in memory and uses that copy instead of reloading the form.
    
2. Forms already cached in the form cache on the local disk drive.
    
3. Forms published in the folder that is currently selected.
    
4. Forms in the Personal Forms Library.
    
5. Forms in the Organizational Forms Library.
    
6. Standard Outlook forms, such as Note, Post, and Contact, in the Application Forms Library.
    

 **Note**  Before it looks for a custom form, Outlook determines whether the message class of any form region matches the message class of the item being loaded. If there is a match, Outlook loads the form region. By default, Outlook also loads any form region that has a message class derived from the message class of the item, unless the  **exactMessageClass** element of the form region is set to **True**. After loading the appropriate form regions, Outlook proceeds to look for forms in the order specified above. However, if any of the loaded form regions is a replacement or replace-all form region that has the  **loadLegacyForm** element set to **False**, Outlook will not continue to look for and load any form that contains custom form pages. For more information about the  **exactMessageClass** and **loadLegacyForm** elements, see the Outlook 2010 XML Schema Reference in the [MSDN Library](http://msdn.microsoft.com/library).

Because Outlook caches forms, avoid having more than one form with the same name or publishing the same form to more than one forms library. Forms that are used in a folder-based solution should be published only in the folder. If you are developing a solution based on mail message forms, you can temporarily publish the forms in your Personal Forms Library. After you finalize a form, publish it to the Organizational Forms Library on the Microsoft Exchange Server. Make a backup copy of the form, and then delete it from your Personal Forms Library. If you need to publish a form in more than one location, make sure that you keep all forms libraries up-to-date with the current version of the form.

 **Note**  


