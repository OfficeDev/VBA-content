---
title: Using the Form Region XML Manifest to Define a Form Region
ms.prod: outlook
ms.assetid: a1c150b1-a6ee-6f16-9798-82d253cbcc6a
ms.date: 06/08/2017
---


# Using the Form Region XML Manifest to Define a Form Region

To run a form region, you must register it in the Windows registry and specify the message class and other necessary information that Microsoft Outlook needs to display the form region. The form region XML schema allows you to specify information related to the functionality, behavior, and other innate properties of the form region. Typically, you would specify XML that follows this schema in a form region manifest XML file, and then register this file in the Windows registry so as to use the form region to display items of the corresponding message class.

For example, if you are designing a form region for items of the message class IPM.Contact, you can specify a form region manifest XML file, contoso.xml, that defines characteristics of the form region. When you register the form region in the Windows registry, under the current user key, you can add the key  **IPM.Contact**, add a value of the type  **String**, and specify the name of the form region,  **ContosoAdjoining**, as the name of the key, and the form region manifest XML file,  **c:\Form Regions\contoso.xml**, as the data of the key. For more information, see  [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).

You can specify the functionality, behavior, and other innate properties of a form region through elements in the form region XML schema. Some of the more commonly used elements are listed as follows. For more information on the form region XML schema, see the Microsoft Outlook 2010 XML Schema Reference in the  [MSDN Library](http://msdn.microsoft.com/library).



| **Schema Elements**| **Purpose**| **Further Information**|
| **name**,  **title**,  **formRegionName**|Identify a form region internally and in the user interface.| [How-to: Name a Form Region](name-a-form-region.md)|
| **formRegionType**|Specify a form region to occupy part of a page or an entire page of a form.| [How to: Specify the Location of a Form Region in a Custom Form](specify-the-location-of-a-form-region-in-a-custom-form.md)|
| **displayAfter**|Order multiple form regions in a custom form.| [How to: Specify the Location of a Form Region in a Custom Form](specify-the-location-of-a-form-region-in-a-custom-form.md)|
| **layoutFile**|Specify a layout file for a form region.| [How to: Specify a Layout File for a Form Region](specify-a-layout-file-for-a-form-region.md)|
| **showInspectorCompose**|Prevent the inspector from displaying a form region when you are composing a message.| [How to: Prevent the Inspector from Displaying a Form Region When You are Composing a Message ](prevent-the-inspector-from-displaying-a-form-region-when-you-are-composing-a-mes.md)|
| **showInspectorRead**|Prevent the inspector from displaying a form region when you are reading a message.| [How to: Prevent the Inspector from Displaying a Form Region When You are Reading a Message](prevent-the-inspector-from-displaying-a-form-region-when-you-are-reading-a-messa.md)|
| **showReadingPane**|Prevent the Reading Pane from displaying a form region when you are previewing a message.| [How to: Prevent the Reading Pane from Displaying a Form Region When You are Previewing a Message](prevent-the-reading-pane-from-displaying-a-form-region-when-you-are-previewing-a.md)|
| **hidden**|Prevent a form region from being modified in the Forms Designer.| [How to: Prevent a Replacement Form Region from Being Used to Create a New Item or from Being Modified in the Forms Designer](prevent-a-replacement-form-region-from-being-used-to-create-a-new-item-or-from-b.md)|
| **exactMessageClass**|Specify a form region is to be used only for items that have exactly the same message class as the form region.| [How to: Specify a Form Region to be Used Only for the Exact Message Class](specify-a-form-region-to-be-used-only-for-the-exact-message-class.md)|
| **action** and its child elements|Modify a standard action that is available to a form region.| [How to: Modify a Built-in Action for a Form Region](modify-a-built-in-action-for-a-form-region.md)|
| **action** and its child elements|Create a custom action for a form region.| [How to: Create a Custom Action for a Form Region](create-a-custom-action-for-a-form-region.md)|
| **icons** and its child elements|Specify custom icons for a form region.| [How to: Specify Icons to be Displayed for a Form Region](specify-icons-to-be-displayed-for-a-form-region.md)|
| **stringOverride** and its child elements|Specify locale-specific strings in the user interface of a form region.| [How to: Specify Locale-Specific User Interface for a Form Region](specify-locale-specific-user-interface-for-a-form-region.md)|

