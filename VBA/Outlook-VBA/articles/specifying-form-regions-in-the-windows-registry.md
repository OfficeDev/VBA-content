---
title: Specifying Form Regions in the Windows Registry
ms.prod: outlook
ms.assetid: 0de3fcb1-b357-8300-c943-9a5a788d4976
ms.date: 06/08/2017
---


# Specifying Form Regions in the Windows Registry

To run a form that contains a form region on a client computer, you must register the form region in the Windows registry, specifying the message class and other information necessary for Microsoft Outlook to display the form region. This information includes the display name, where the form region appears in a form, any layout file or add-in that implements the form region, any supported user actions, and any localized terms for the user interface. The structure of this information follows a form region XML schema; for more information on the XML schema for form regions, see the Microsoft Outlook 2010 XML Schema Reference in the  [MSDN Library](http://msdn.microsoft.com/library). 

There are a few ways to specify this information about the form region in the registry. You can explicitly specify the XML, or a full path to an XML file, that contains this information for the form region and that conforms to the form region XML schema. You can also choose to specify the ProgID of an add-in which will provide Outlook the XML manifest for the form region. When Outlook starts, it reads the list of form regions from the registry and caches the associated data.

 **Caution**  Incorrectly editing the Windows registry may severely damage your system. Before making changes to the registry, you should back up any valued data on the computer.


## Registering a Form Region

 Register form regions under the **FormRegions** key in the Windows registry, under the local machine key (as **HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\FormRegions**) or under the current user key (as  **HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\FormRegions**). Under the  **FormRegions** key, create a separate key for each message class for which form regions have been created. For example, the mail item has the message class **IPM.Note**; you can register all form regions used to display the mail item for the current user under the key,  **HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\FormRegions\IPM.Note**.

 The following steps show how to register a form region under the local machine:


1. Close Outlook.
    
2. Add the following key to the registry if it does not already exist:  **HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\FormRegions**.
    
3. Under the  **FormRegions** key, add a key with the name of the message class that the form region is associated with, if that key does not already exist. For example, to create a form region for the contact item, add a key with the name **IPM.Contact**, if it does not already exist.
    
4. For that key, add a value of the type  **REG_SZ**, and specify the name to be the same as the  **name** element of the form region. This is the internal name for the form region; the internal name supports only ASCII characters. Specify the data as one of the following possible values:
    
      - Explicitly the XML that specifies the layout, behavior, and other characteristics for the form region, and that conforms to the form region XML schema. In this case, you must precede the data with a less than sign ( **&lt;**).
    
  - The ProgID of an add-in that will provide Outlook the XML manifest for the form region. In this case, you must precede the data with an equal sign ( **=**). For example, if the ProgID of an add-in is MyAddinProject1.ConnectClass, you would specify the data of the key to be  **=MyAddinProject1.ConnectClass**.
    
  - The full local file path name to a form region XML manifest file that describes the layout, behavior, and other characteristics of the form region. If the data does not precede with a less than sign ( **&lt;**) or an equal sign ( **=**), then Outlook will assume the data is a path name to the form regions XML manifest file. For example, if your form region XML manifest file, map.xml, is at c:\Form Regions\, you would specify the data of the key to be  **c:\Form Regions\map.xml**.
    



## Specifying Form Regions as Replacements for Standard Forms

Outlook allows you to customize the standard form for each of the Outlook message classes by adding adjoining form regions or separate form regions to the form. The following table shows the standard forms and corresponding message classes in Outlook. 



| **Standard Form**| **Message Class**|
|Appointment| **IPM.Appointment**|
|Contact| **IPM.Contact**|
|Journal Entry| **IPM.Activity**|
|Meeting Request| **IPM.Meeting.Schedule.Request**|
|Message| **IPM.Note**|
|Post| **IPM.Post**|
|Task| **IPM.Request**|
|Task Request| **IPM.Task**|

 **Note**  You cannot specify form regions for the root Outlook message class,  **IPM**. 

You can add separate form regions as extra pages to a standard form, but you cannot replace any existing pages on the standard form and keep the form for the same Outlook message class. If you need to replace the default page or all pages of a standard form, you will have to derive a new message class for that form, specify a replacement form region to replace the default page or a replace-all form region to replace the entire form, and register that form region for the derived message class.

For example, you can create a replacement form region that replaces the  **General** page of the Contact form, and register that form region for a message class derived from **IPM.Contact**, such as  **IPM.Contact.MyContact**. You cannot register the form region for the  **IPM.Contact** message class.

 When Outlook opens an item and sees a derived message class (for example, **IPM.Contact.MyContact.Personal.Family**), it looks for a replacement or replace-all form region (that is, a form region that has a  **formRegionType** element being equal to **replace** or **replaceAll**), and that matches exactly the derived message class,  **IPM.Contact.MyContact.Personal.Family**. If there is no exact match, Outlook will try  **IPM.Contact.MyContact.Personal**, and if that fails, Outlook will try  **IPM.Contact.MyContact**. Note that Outlook ignores any replacement or replace-all form regions for  **IPM.Contact**. If there is not yet an exact match, Outlook will look for any form region with  **formRegionType** equal to **adjoining** or **separate** for the derived class, **IPM.Contact.MyContact.Personal.Family**.


## Multiple Form Regions for the Same Message Class

When one or more add-ins registers multiple form regions for the same message class, the display order of adjoining form regions on the default page and the order of separate form regions in the form depends on the order that the add-ins are installed and the order that the add-ins register the form regions. If an add-in specifies more than one adjoining form region or more than one separate form region for a message class, the add-in can use the  **displayAfter** element to specify the order of these form regions. The order specified by **displayAfter** takes precedence over the order of the form regions in the registry. This is the only means add-ins can specify the order of form regions in a form.

After the form regions are installed on a client computer, form users can further customize the order of adjoining form regions by opening the form and moving the form regions up or down on the default page through the form region header context menu.


## Example

The following is an example of the XML for a form region for a derived message class,  **IPM.Contact.MapUser**. The form region is applied to all users on a computer. The XML file, map.xml, is located in c:\Form Regions.

To register the form region, create the following key in the Windows registry:



| **Key**|HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\FormRegions\IPM.Contact.MapUser|
| **Name**|MapTab|
| **Type**|REG_SZ|
| **Data**|c:\Form Regions\map.xml|


The following lists the content of map.xml: 




```xml
<?xml version="1.0"?> 
<FormRegion xmlns="http://schemas.microsoft.com/office/outlook/12/formregion.xsd">   
    <!-- Internal name --> 
    <name>MapTab</name> 
    <!-- Display name --> 
    <title>Directions</title> 
    <!--  Additive separate form region --> 
    <formRegionType>separate</formRegionType> 
    <!--  Layout file --> 
    <layoutFile>Map.ofs</layoutFile> 
    <!-- Icon for form region in all contexts --> 
    <icons> 
        <default>generic.ico</default> 
    </icons> 
</FormRegion> 
```

The form region is added to the form as a page following the last non-hidden built-in page in the Contact form (normally, this would follow the  **All Fields** page). The page is titled **Directions** and has an internal programmatic name "MapTab". Map.xml specifies a layout file and an icon file. Note that all file paths in the xml file can be specified as full file paths, or paths relative to the location of the form region XML manifest file.


