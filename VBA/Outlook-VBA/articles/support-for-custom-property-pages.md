---
title: Support for Custom Property Pages
keywords: vbaol11.chm5268752
f1_keywords:
- vbaol11.chm5268752
ms.prod: outlook
ms.assetid: a2d98281-486b-2f91-8479-080951c3e663
ms.date: 06/08/2017
---


# Support for Custom Property Pages

You can create your own property pages to customize the  **Properties** dialog box for all folders.

A custom property page is an ActiveX control stored in a dynamic-link library (DLL) that implements the  [PropertyPage](propertypage-object-outlook.md) object and that's installed as a [COM add-in](support-for-com-add-ins.md) . This object provides the interface through which Outlook can query the property page about its status and inform the property page that the user has clicked the **Apply** or **OK** button.

For more information about property pages, see  [adding custom property pages](adding-custom-property-pages.md).


 **Note**  Customizing the  **Outlook Options** dialog box (available through the Microsoft Office Backstage view) by using property pages has been deprecated. However, you can customize your own tab on the Backstage view using Microsoft Office Fluent user interface extensibility. For more information, see [Extending the User Interface in Outlook 2010](http://msdn.microsoft.com/library/00b504b0-e897-43b9-8615-44276166823f.aspx).


