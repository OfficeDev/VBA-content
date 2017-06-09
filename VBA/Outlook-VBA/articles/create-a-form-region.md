---
title: Create a Form Region
ms.prod: outlook
ms.assetid: 695b95a5-c795-cb4a-8d35-ba12b0007b1f
ms.date: 06/08/2017
---


# Create a Form Region

This topic lists the considerations and steps for creating a form region. 

There are different types of form regions, depending on how you want to customize the form and where you place the form region in the form. A form region can add extra user interface to the default page or add an extra page to a standard form. Or, it can replace the default page of a standard form, or replace the entire standard form, resulting in a new form for a derived message class. You specify the type of the form region in a form region manifest XML file, using the <formRegionType> tag as described in step 7. Aside from this, the means to create and design these various types of form regions is identical:

1. Use the Forms Designer to create and design the layout (steps 1 through 5)
    
2. Save the form region to an Outlook Form Storage (.OFS) file (step 6)
    
3. Create a form region manifest XML file to specify other details about the form region to Microsoft Outlook (step 7)
    
4. Register the form region for a message class (step 9)
    
When you restart Outlook, the form region will be available for use. Alternatively, you can create the form region and the form region manifest XML file, use an add-in to extend the form region and register the form region programmatically. When you install the add-in, the add-in will also install the form that contains the form region. For more information, see  [Extending a Form Region with an Add-in](extending-a-form-region-with-an-add-in.md).
The following procedure details the steps to create a form region using the Forms Designer without an add-in.

1. On the  **Developer** tab, in the **Custom Forms** group, click **Design a Form**.
    
2. Select one of the nine standard Outlook forms that is best suited for your purpose:  **Appointment**,  **Contact**,  **Journal Entry**,  **Meeting Request**,  **Message**,  **Post**,  **RSS Article**,  **Task**, and  **Task Request**. 
    
    When you customize a form, you always start with a standard form as a template. When you choose the standard form, you should consider the following: 
    
      - The actions associated with the form, for example, whether you will be sending the form to other people, in which case you should choose the Message form. 
    
  - The kinds of fields you will need in the form, for example, whether they are mostly contact-specific fields.
    
3. Plan the scope of your customization. Will it suffice if you add extra controls to the bottom of the default page of the standard form? Will you need an extra custom page? Or, will you want to modify the user interface so substantially that it would be easier to create a new form? Note that you can replace pages on a form only if you specify that form for a derived message class.
    
4. In the Forms Designer, on the  **Developer** tab, in the **Design** group, click **New Form Region**.
    
    Note that any type of form region opens as a separate page in the Forms Designer. At runtime, the form region will be displayed the way you have specified in the form region manifest XML file, as described in step 7.
    
5. Design the layout of the form region by dragging and dropping controls from the Toolbox to the form region, and binding them to fields where appropriate.
    
    Similar to customizing a form page, customizing a form region involves defining custom fields, inserting controls using the Toolbox, and binding the controls with fields using the Field Chooser. For more information, see  [Controls in a Custom Form](controls-in-a-custom-form.md) and the section Design the Form Region in [Walkthrough: Add a Form Region to an Existing Page on a Form](add-a-form-region-to-an-existing-page-on-a-form.md). Optionally, you can use an add-in to program events of the controls.
    
6. Save the layout of the form region by clicking  **Save Region**, and then  **Save Form Region** in the **Design** group. The form region layout file will be saved with a .OFS extension.
    
7. Use an XML editor such as Notepad to create a form region manifest XML file.
    
    You must specify XML for each form region to tell Outlook how to display it and the actions it supports. The XML must validate against the form region manifest schema (for more information about the schema, see the Microsoft Outlook 2010 XML Schema Reference in the  [MSDN Library](http://msdn.microsoft.com/library)). The schema supports many elements, such as the more commonly used ones enumerated below:
    
      - The \<addin\> tag specifies the ProgID of the add-in that manages the form region and provides storage for it. You should only specify this tag if you use an add-in to create and manage the form region.
    
  - The \<customActions\> tag that specifies the actions supported by the form region, for example, reply and forward.
    
  - The \<displayAfter\> tag specifies the form region that precedes the current form region in the same form. This information defines the initial order of multiple adjoining form regions or multiple separate form regions in the same form. 
    
  - The \<formRegionType\> tag specifies whether the form region is an additive form region (adjoining or separate form region), or is a replacement or replace-all form region (replacing the default page or the entire standard form and resulting as a new form for a derived message class).
    
  - The \<layoutFile\> tag specifies the location of the .OFS file, if one exists. Note that any file paths in the .OFS file, including this file path, can be specified as a relative path to the location of the form region manifest XML file specified in the registry. However, also note that UNC paths are not supported. If you use an add-in to create and manage the form region, you must specify the \<addin\> tag but not this tag.
    
  - The \<name\> tag specifies the name for the form region used only in code.
    
  - The \<title\> tag specifies the display name of a separate form region in the  **Actions** menu and the **Choose Form** dialog box.
    
  - The \<icons\> tag specifies the location of icon files. 
    
     **Note**  By default, the icon file is in the same folder as, or in a path relative to, the form region manifest XML file. You can also specify a full path for the icon file, for example: `<icons><default>c:\myicon.ico</default></icons>`or a full path for a resource file, for example: `<icons><unread>c:\myresource.dll,101</unread> </icons>`which loads the icon resource 101 in the resource file c:\myresource.dll. However, do not use the implicit convention that specifies icons embedded in the add-in assembly file. For example: `<icons><read>,102</read></icons>`will not be supported and will not load the icon resource 102 in the add-in dll.
8. Close Outlook.
    
9. Register the form region in the Windows registry, specifying the message class that this form region is intended for, and the full file path for the form region manifest XML file.
    
    Register form regions under either the  **HKEY_CURRENT_USER** or the **HKEY_LOCAL_MACHINE** hive in the Windows registry. For example, additive form regions for the **IPM.Contact** message class for the current user should be registered under the same key, **HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\FormRegions\IPM.Contact**. Note that the form region will be displayed for the current user in all forms applied to  **IPM.Contact** and any message class derived from **IPM.Contact**. If you want a form region to be used only for  **IPM.Contact** and do not wish derived message classes to use that form region, you can specify this using the \<exactMessageClass\> tag in the form region manifest XML file.
    
10. Start Outlook. When you open an item of the message class that you have specified for the form region in step 9, you will see the form region in the inspector.
    

