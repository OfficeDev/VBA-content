---
title: Save a Form with the Item (One-off Forms)
ms.prod: outlook
ms.assetid: 0983163c-6ae0-a391-ae31-afd7ec796d4b
ms.date: 06/08/2017
---


# Save a Form with the Item (One-off Forms)

 **Note**  This Help topic describes how to save a form that is customized with form pages; it does not apply to forms that are customized with form regions. With form regions, you save the layout of the form region by clicking  **Form Region**, and then  **Save Form Region** to save the form region layout file with an .ofs extension. After that, you must create a form region manifest XML file and register the form region in the Windows registry. For more information, see [How to: Create a Form Region](create-a-form-region.md).


## Forms customized with form pages

When you create a solution by using custom forms in Microsoft Outlook, it is important to understand how Outlook uses your custom form in relation to the items in a folder.

For example, when you create a custom contact form to replace the default Outlook contact form, you typically follow these steps:


1. Start with a new, default contact item to use as the basis for your custom form.
    
2. Modify the form to meet your needs.
    
3. Publish the form to the  **Contacts** folder.
    
4. Set the form as the default form for the  **Contacts** folder by changing the folder properties.
    
In this typical scenario, information about the form (the form definition) is not saved with each item. Instead, the form is stored in the location where it was published and is referenced by using the  **Message Class** field. This way, each item only stores the data associated with it, and its size is relatively small.

However, it is possible to have Outlook store the form definition within individual items in a folder. These items are called one-off items and always use the form definition that is stored within the item instead of the published form.

In most situations, the form definition should not be stored within the item. The most common exception to this is a custom e-mail message form. If you use Microsoft Exchange Server, you can publish a custom e-mail message form to the Organizational Forms Library so that it is always available to everyone in the organization. This way, you do not have to store the form definition in the item. If you do not use Exchange Server, or if you are sending the form to another organization where the form is not available, select the  **Send form definition with item** check box on the **Properties** page of the form when in design mode. Depending on security restrictions, this might enable the recipient to view the e-mail message with the custom form.


 **Note**  If the recipient still cannot view your custom form, make sure that you customized or disabled the  **Read** page of the custom e-mail form.

If the custom form contains Microsoft Visual Basic Scripting Edition (VBScript), Outlook displays the macro virus warning unless the form is published in the Exchange Server Organizational Forms Library.

The following scenarios commonly result in items becoming one-off items.


- You have a folder-based solution whereby the form is published in the folder and the items use the published form. You open an existing item in a folder, make changes to the form in design mode, and then save the item.
    
    Because the form definition has changed and the form was not republished, Outlook saves the new form definition with the item. To change the form for all items in the folder, instead of opening an existing item, follow these steps:
    
      1. Open a new item based on your custom form.
    
  2. Make the form design changes to that item.
    
  3. Republish the form with the same name.
    
  4. Close and do not save changes to the item.
    

    All the items in the folder will use the updated custom form the next time that the items are opened, because the message class of the items still refers to the published form.
    
- VBScript code in the custom form changed the form definition of the item.
    
    If VBScript code within an item programmatically changes the form, in many cases the result is that the form definition is saved with the item. The following Outlook object model methods most commonly cause this behavior:
    
      -  **UserProperties.Add** method.
    
  - Methods and properties of the  **[FormDescription](formdescription-object-outlook.md)** object.
    
  - Some methods or properties of controls, such as  **Enabled**.
    
  - Methods and properties of the  **[Actions](actions-object-outlook.md)** collection object.
    
Although solutions and situations vary greatly, the following might indicate that an item has become a one-off item.


- VBScript code in the form does not run, or a macro virus warning unexpectedly appears, indicating that the item itself, and not just a published form, contains VBScript code.
    
- The size of an item increases unexpectedly.
    
- An item's icon changes unexpectedly.
    

