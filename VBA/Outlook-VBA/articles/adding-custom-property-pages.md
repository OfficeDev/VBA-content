---
title: Adding Custom Property Pages
keywords: vbaol11.chm5272720
f1_keywords:
- vbaol11.chm5272720
ms.prod: outlook
ms.assetid: 45390e9b-438c-86b0-488c-b179cabe4276
ms.date: 06/08/2017
---


# Adding Custom Property Pages

Creating a custom property page for Microsoft Outlook involves four major steps:


1. Create the page as an ActiveX control.
    
2. Implement the  [PropertyPage](propertypage-object-outlook.md) object.
    
3. Write a procedure that sets the value of the  [Dirty](propertypage-dirty-property-outlook.md) property and calls the [OnStatusChange](propertypagesite-onstatuschange-method-outlook.md) method.
    
4. Create a Component Object Model (COM) add-in that contains an event procedure for  **OptionsPagesAdd** for either the [Application](application-object-outlook.md) or [NameSpace](namespace-object-outlook.md) object, as appropriate. For information about creating a COM add-in, see [Customizing Outlook using COM add-ins](customizing-outlook-using-com-add-ins.md).
    

## Create the page as an ActiveX control

A custom property page in Outlook is an ActiveX control that's implemented along with a dynamic link library (DLL) that's designed as a COM add-in. The easiest way to create a custom property page is using Microsoft Visual Basic version 6.0 or higher. This version of Visual Basic provides templates and tools that simplify the process of creating both ActiveX controls and COM add-ins.

When you create the ActiveX control, you populate it with the controls your user will require to set the properties your page is designed to support. Because Outlook might resize the control when it displays the property page in the dialog box, the control's  **Initialize** event should position and size the child controls dynamically, depending on the final values of the control's **Width** and **Height** properties.

The dialog box in which the custom property page is displayed has three buttons below the property pages: an  **OK** button, a **Cancel** button, and an **Apply** button. When the user clicks the **OK** button, changes to properties on all pages in the dialog box are applied and the dialog box is closed. If the user clicks the **Cancel** button, no properties are changed and the dialog box is closed. If the user clicks the **Apply** button, any changes to properties are applied but the dialog box remains open. You should design your property page to respond appropriately when the user clicks these buttons. Later sections in this topic describe how to notify Outlook that the status of your property page has changed and how Outlook notifies your program when the changed property values should be applied.


## Implement the PropertyPage object

The  **PropertyPage** object is an abstract object; that is, its interfaces are defined but not implemented by Outlook. If your custom property change will rely on the **Apply** or **Help** button of the parent dialog box, the module that contains the custom property page ActiveX control must implement the **PropertyPage** object. To implement the object, the module must have a reference set to the Outlook object library and must contain the following statement.


```vb
Implements Outlook.PropertyPage
```

The module must then contain code that implements the methods and properties of the  **PropertyPage** object. The following table describes these procedures.



|**Procedure**|**Description**|
|:-----|:-----|
| **Dirty** property|Called by Outlook in response to the  **OnStatusChange** method to determine whether the user has changed a value on the property page.|
| [Apply](propertypage-apply-method-outlook.md) method|Called by Outlook to notify your program that the user has clicked the  **OK** or the **Apply** button. Usually this procedure applies any property values changed by the user in the property page.|
| [GetPageInfo](propertypage-getpageinfo-method-outlook.md) method|Called by Outlook to obtain the Help file and topic associated with the property page.|

## Write a procedure that sets the Dirty property and calls the OnStatusChange method

Most commonly, changes to property values are not applied immediately in response to user interaction with the controls that let the user specify those values. Instead, the values are applied when the user clicks  **OK** or **Apply** in the dialog box. The **Apply** button is grayed until the user changes a value on a property page. To notify Outlook that the user has changed a value on your property page, your program should call the **OnStatusChange** method and then return True when Outlook queries the Dirty property.


## Create a COM add-in containing an event procedure for the OptionsPagesAdd event

The  **OptionsPagesAdd** event gives your program the opportunity to add your custom property page to the Microsoft Outlook **Options** dialog box (if the event is fired for the **Application** object) or the folders **Properties** dialog box (if the event is called for the **NameSpace** object). When Outlook calls this event procedure, it passes a [PropertyPages](propertypages-object-outlook.md) object. Your event procedure uses the [Add](propertypages-add-method-outlook.md) method of the collection to add the **PropertyPage** object implemented by your program to the object.


