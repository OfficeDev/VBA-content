---
title: Specify Icons to be Displayed for a Form Region
ms.prod: outlook
ms.assetid: 9ffb9f46-a3b9-d90c-6771-9cd9f9b2e04a
ms.date: 06/08/2017
---


# Specify Icons to be Displayed for a Form Region

When you define a form region for a custom message class, such as  **IPM.Note.Customer**, you can specify custom icons to be displayed in the explorer, inspector, and the ribbon for items belonging to that message class. 

Icons are specified as resources in a Win32 resource DLL file. You can refer to an icon file in the standard Win32 embedded icon notation. For example:

-  `<default>c:\myicon.ico<.default>` loads the default icon from the specified path, c:\
    
-  `<unread>c:\myresource.dll,101</unread>` loads the icon resource that has the resource ID 101 from the resource file myresource.dll in the specified path, c:\
    

Note that you can specify environment variables in the file path names, but you cannot specify paths in Universal Naming Convention (UNC).
By default, if you do not specify any custom icons, the icons assigned to the parent message class will be used. For example, if you do not specify any icons for a form region for  **IPM.Note.Customer**, then the icons for  **IPM.Note** will be used.
Depending on the type of item, there are different states of the item that you can consider distinguishing with separate icons. For example, in the explorer, a task item can use a custom icon to identify itself as recurrent, and a mail item can use a custom icon to identify itself as having been replied to. You do not have to specify a separate icon for each state that the type of item can be in; you can choose to specify a default icon that will be displayed in all states in the explorer, inspector, and ribbon that apply for that item type.
The following table shows the states of an item that you can consider to identify with custom icons in the explorer, inspector, or ribbon. All custom icons for a form region are specified under the  **icons** element in the form region manifest XML file for that form region. Each state is mapped with an XML child element of the **icons** element. You will specify this form region manifest XML file when you register the form region in the Windows registry. For more information on registering a form region, see [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).


| **State of an Item**| **XML Child Element**| **Example**|
|Any state that applies to the item, if no other custom icon has been defined for that state.| **default**| `<default>c:\default.ico</default>`|
|Icon to identify in the explorer that item has been encrypted.| **encrypted**| `<encrypted>c:\encryptedicon.ico</encrypted>`|
|Icon to identify in the explorer that item has been forwarded.| **forwarded**| `<forwarded>c:\forwardedicon.ico</forwarded>`|
|Icon to identify in the ribbon that item belongs to a specific derived message class. | **page**| `<page>c:\pageicon.ico</page>`|
|Icon to identify in the explorer that item has been read.| **read**| `<read>c:\readicon.ico</read>`|
|Icon to identify in the explorer that item is recurrent.| **recurring**| `<recurring>c:\recurringicon.ico</recurring>`|
|Icon to identify in the explorer that item has been replied to.| **replied**| `<replied>c:\repliedicon.ico</replied>`|
|Icon to identify in the explorer that item has been signed with a digital signature.| **signed**| `<signed>c:\signedicon.ico</signed>`|
|Icon to identify in the explorer that item has been sent.| **submitted**| `<submitted>c:\submittedicon.ico</submitted>`|
|Icon to identify in the explorer that item has not yet been read.| **unread**| `<unread>c:\unreadicon.ico</unread>`|
|Icon to identify in the explorer that item is pending and has not yet been sent.| **unsent**| `<unsent>c:\unsenticon.ico</unsent>`|
|Icon to be displayed in the inspector when this item type has been opened.| **window**| `<window>c:\windowicon.ico</window>`|

## To specify a custom icon for a form region


1. In the form region manifest XML file, under the  **icons** element, specify the child element that maps to the state that you would like to customize.
    
2. Depending on how you would like the custom icon file to be specified, do either of the following: 
    
      - If you want Outlook to load the icon from an icon file or a resource file, specify the location of the icon file or resource file in the child element.
    
  - If you want an add-in to inform Outlook which icon to display, specify  `addin` in the child element.
    
The following example specifies custom icons for several states of an item belonging to the message class supported by a form region: 


```XML
<icons> 
 <default>c:\icons\MyIcon.ico</default> 
 <unread>c:\icons\MyUnReadIcon.ico</unread> 
 <read>c:\icons\MyReadIcon.ico</read> 
 <encrypted>%windir%\myresource.dll,101</encrypted> 
</icons>
```

The four custom icons include:


- A custom icon file for the read state
    
- A custom icon file for the unread state
    
- A location in a resource file for the encrypted state
    
- A default icon file for all other states applicable to the item
    

 **Note**  The value of the child element can be expressed either as a file path to an icon file or a resource file, or as  `addin`. The file path can be expressed as a full path or a path relative to the location of the form region manifest XML file, and can involve system variables. For more information on specifying an icon using an add-in, see  [How to: Use an Add-in to Specify Icons for a Form Region](use-an-add-in-to-specify-icons-for-a-form-region.md).


