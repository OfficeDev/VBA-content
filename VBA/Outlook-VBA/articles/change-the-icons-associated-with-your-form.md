---
title: Change the Icons Associated with your Form
ms.prod: outlook
ms.assetid: 2c13a1ad-b901-cda4-9f14-ce3357ab98e3
ms.date: 06/08/2017
---


# Change the Icons Associated with your Form




## Forms customized with form regions

When you customize icons on a form region, you can set icons for each of the different states of the form as well, such as replied, forwarded, or unread. For more information, see  [How to: Specify Icons to be Displayed for a Form Region](specify-icons-to-be-displayed-for-a-form-region.md).


## Forms customized with form pages


1. In the Forms Designer, click the  **Properties** page.
    
2. Click  **Change Large Icon** or **Change Small Icon**. 
    
3. Select the icon that you want to use and then click  **Open**.
    

|**Note**|
|:-----|  
|<p>Sometimes you may find that even if you specify that a custom form should use a particular icon, when you view an item in a folder, a standard icon is used instead. This can happen in the following scenarios: </p><ul><li><p>You published the form in a location that is not accessible to everyone. For example, when they receive a custom e-mail message form, they do not have access to the custom form. Therefore, Outlook displays the standard icon.</p><p>**Solution** Publish the form to a location that is accessible to everyone who uses the form, typically the Organizational Forms Library.</p></li><li><p>The item has become a one-off form. In this case, the message class of the item changes to the default message class for that particular type of item, and the icon reverts back to the default icon for that type of item.</p><p>**Solution** Reset the message class of the item.</p></li><li><p>You replied to or forwarded an item in a folder, and the new item does not use the custom icon. In this case, the form action that started the new item did not specify a custom message class. Therefore, the new item is a standard form that does not have a custom icon.</p><p>**Solution** Disable the standard Reply, ReplyAll, and Forward actions and create custom Reply, ReplyAll, or Forward actions that start your custom form instead.</p></li><li><p>You replied to or forwarded an item in a folder. In this case, Outlook replaces the custom icon with the standard icon so that the reply or forward indicator arrows can be displayed.</p><p>**Solution** This is a design limitation of Outlook. The forward or reply icons are always used instead of your custom icon.</p></li></ul>|