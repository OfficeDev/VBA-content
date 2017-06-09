---
title: Overview of the IRibbonControl Object
ms.prod: outlook
ms.assetid: 32a0ae0b-26d9-673b-d609-b86696538435
ms.date: 06/08/2017
---


# Overview of the IRibbonControl Object

The  [IRibbonControl](http://msdn.microsoft.com/library/63aef709-e1d3-b1a6-76af-b568ad0e69ae%28Office.15%29.aspx) object is passed in most of the callbacks that are available for controls in the ribbon or Microsoft Office Backstage view, as well as the customizable menu items in Microsoft Outlook. The object is especially useful for Outlook developers because it provides an [IRibbonControl.Context](http://msdn.microsoft.com/library/39f9d85a-00e9-9682-3957-51d9e72b4d83%28Office.15%29.aspx) property that returns the related Outlook object to which the customization is applied and is about to be displayed. 

For example, the **Context** property returns the [Explorer](explorer-object-outlook.md) object if you customize the ribbon in an explorer, and returns the [Store](store-object-outlook.md) object if you customize the shortcut menu for a store folder.

 **IRibbonControl** exposes the following properties.


| **Property**| **Type**| **Description**|
|:-----|:-----|:-----|
| **Context**| **Object**| Returns an object that represents the window in which the custom ribbon is about to be displayed, or the related object to which the menu customization is applied and is about to be displayed. Read-only.|
| **[Id](http://msdn.microsoft.com/library/56a0d143-66de-ab77-0c21-d34341ce5da4%28Office.15%29.aspx)**| **String**|Returns a string that represents the  **Id** attribute for the control or custom menu item. Read-only.|
| **[Tag](http://msdn.microsoft.com/library/d0f041c0-d7bc-7a4f-df9b-ba62fa08f1ca%28Office.15%29.aspx)**| **String**|Returns a string that represents the  **Tag** attribute for the control or custom menu item. Read-only.|
When you write managed code, try to cast the object represented by  **IRibbonControl.Context** to the corresponding Outlook object. For example, if you customize a ribbon in an inspector, cast the [Inspector](inspector-object-outlook.md) object. Then, if the cast succeeds, you can compare the **Inspector** object that is returned by **IRibbonControl.Context** to other inspector windows that are open. To determine the underlying item that is displayed in an inspector window, examine [Inspector.CurrentItem](inspector-currentitem-property-outlook.md). Because  **CurrentItem** is an **Object** type, your code must cast the object to an appropriate item type such as [MailItem](mailitem-object-outlook.md) or [ContactItem](contactitem-object-outlook.md).

## See also


#### Concepts


 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)<br>
 [Implementing the IRibbonExtensibility Interface](implementing-the-iribbonextensibility-interface.md)<br>

