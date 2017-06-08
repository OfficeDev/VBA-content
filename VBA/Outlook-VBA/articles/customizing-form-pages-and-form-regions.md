---
title: Customizing Form Pages and Form Regions
ms.prod: outlook
ms.assetid: c8c2d080-66a8-b761-bdc0-527b209e0bd1
ms.date: 06/08/2017
---


# Customizing Form Pages and Form Regions

You can customize a form in Outlook in two ways: by customizing pages of the form, and by creating form regions for the form. Customizing form pages has been supported since Microsoft Office Outlook 97, and form regions are supported since Microsoft Office Outlook 2007.

This topic compares customizing forms on form pages with customizing forms on form regions, and identifies the advantages of form regions.

|**Note**|
|:-----|  
|Even though custom form pages and form regions can exist in the same custom form, for simpler deployment and easier maintenance, you should develop future custom form solutions using only form regions. For more information on your options for existing form solutions, see  [Best Practices to Migrate Outlook 97-2003 Custom Forms](best-practices-to-migrate-outlook-97-2003-custom-forms.md).|


| **Comparison Aspect**| **Form Pages**| **Form Regions**|
|:-----|:-----|:-----|
| **Outlook version**|Supported since Microsoft Office Outlook 97.|Supported since Office Outlook 2007.|
| **Customization venue**|Adding fields and controls in the Outlook forms designer. Optionally programming controls in VBScript using the Script Editor. |Adding fields and controls in the Outlook forms designer. Optionally programming controls using an add-in.|
| **Controls**|When running on versions of Outlook prior to Office Outlook 2007, form pages support Microsoft Forms 2.0 controls and some third-party ActiveX controls, but not Outlook controls. When running on Office Outlook 2007 or later, form pages support Forms 2.0 controls and Outlook controls. Form pages always display Forms 2.0 controls with a classic look. For more information, see  [Controls in a Custom Form](controls-in-a-custom-form.md).|Form regions support Microsoft Forms 2.0 controls, some third-party ActiveX controls, and Outlook controls. By default, Outlook replaces some Forms 2.0 controls that have themed Outlook counterpart controls by the corresponding themed controls, and therefore always displays them with a themed look. For more information, see  [Controls in a Custom Form](controls-in-a-custom-form.md).|
| **Scope of customization in design time**|You can only customize the following pages: <ul><li><p>The <span class="ui">Message</span> page of the message form  (<b>IPM.Note</b>)</p></li><li><p>The <span class="ui">Message</span> page of the post form (<b>IPM.Post</b>)</p></li><li><p>The <span class="ui">General</span> page of the contact form (<b>IPM.Contact</b>)</p></li><li><p>Up to 5 form pages (P2 to P6) for each standard Outlook form</p></li></ul> You can hide supplementary pages in a standard form (for example, the **Details** page of the contact form), but you cannot customize them.|You can customize the following pages:<ul><li><p>The default page of any standard form.</p></li><li><p>Up to 30 separate  form regions as extra pages.</p></li></ul> You can also use a standard form as a template, replace the default page by a replacement form region or replace the entire standard form by a replace-all form region, and register that form region for a message class derived from the original message class.|
| **Adding user interface to the default page**|You can only add user interface to the default page of the message form, post form, and contact form, and pages P2 through P6 of any standard form. To add user interface to any other default page, for example, the default page of an appointment form, you will need to uncheck  **Display This Page** to hide the page, re-create the default page on a supplementary page like P2, and add custom user interface to that page.|You can add user interface as adjoining form regions to the default page of any standard form.|
| **Adding extra pages**|Up to 5 form pages (P2 to P6) for each form. Check  **Display This Page** to show the page in a form.|Up to 30 separate form regions and 50 adjoining form regions for each form.|
| **Removing default user interface**|You can remove default user interface on only the following pages:<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The <span class="ui">Message</span> page of the message form</p></li><li><p>The <span class="ui">Message</span> page of the post form </p></li><li><p>The <span class="ui">General</span> page of the contact form </p></li></ul> Alternatively, you can hide any default page or supplementary page in a standard form.|You can remove or hide default user interface the same way as under "Form Pages". Alternatively, you can create a replacement form region to "replace" the default page of a form, or create a replace-all form region to "replace" all pages of a form, without having to remove the user interface on the default page. Note that replacement and replace-all form regions are only supported for a custom message class derived from an Outlook message class.|
| **Customizing entire page**|You can hide any page in a standard form, and add custom user interface to pages P2 through P6.|You can hide any page in a standard form, use a separate form region to add an extra page to the form, use a replacement form region to replace the default page of the form, or use a replace-all form region to replace the entire form. Note that any replacement is only supported for a custom message class derived from an Outlook message class.|
| **Support for new (derived) message classes**|The administrator uses the Forms Administrator tool to register a custom form for a derived message class.|You can register form regions for a derived message class in the Windows registry.|
| **Deployment**|The administrator registers and installs the custom form. If an add-in exists for the form, the administrator installs it separately from the form.|An administrator installs the add-in. In turn, the add-in installs files for the form regions and registers the form regions for the custom form.|
| **Display of customization in runtime**|Customization is only displayed in the inspector.|Customization is displayed in the inspector and the Reading Pane.|
| **Support for localized user interface**|No|Supports localized strings for form region names, control names, and user actions based on locale.|
| **Sharing between solutions**| Each custom form can only be customized by one add-in (except through the **[ModifiedFormPages](inspector-modifiedformpages-property-outlook.md)** property of the **[Inspector](inspector-object-outlook.md)** object).|Except for the message class IPM, a form for any message class can be customized by more than one add-in. |


