---
title: Modify a Built-in Action for a Form Region
ms.prod: outlook
ms.assetid: c2493139-5c76-6f1c-6cee-7e0907d94c70
ms.date: 06/08/2017
---


# Modify a Built-in Action for a Form Region

By default, there are four built-in actions available to a form region: reply, reply-all, reply-to-folder, and forward. You can modify a built-in action in the following ways, by making the specifications in the form region manifest XML file that you register for the form region:


- The message class of the item that results from executing the action.
    
- The way that the body of the current item is included in the resulting item.
    
- The way that Outlook responds when executing the action.
    
- The prefix value in the subject of the resulting item.
    
- Disabling the built-in action for the form region.
    

For more information on registering a form region, see  [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).

If customizing a built-in action in the above ways does not suit your needs, then you should consider creating a new custom action. For more information, see  [How to: Create a Custom Action for a Form Region](create-a-custom-action-for-a-form-region.md)

## Identifying Actions for a Form Region

Whether you choose to modify a built-in action or create a custom action, you define the action in the form region manifest XML file. Define these actions under the  **customActions** element, enclosing each action in its own **action** element and identifying it by the mandatory **name** attribute. The value of the **name** attribute is a string that represents the internal name of the action.


### To specify the internal name of a built-in action


1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the built-in action as the value of the  **name** attribute of the **action** element.
    
The following example identifies the two built-in actions,  `reply` and `replyAll`, before modifying them:


```
<customActions>
    <action name="reply">
        <!-- further modify this action -->
    </action>
    <action name="replyAll">
        <!-- further modify this action -->
    </action>
</customActions>
```

Note that by default, there are four built-in actions for each form region. You can identify them with the following keywords:



| **Keyword**| **Built-in Action**|
| **forward**|Forward current item.|
| **reply**|Reply to current item.|
| **replyAll**|Reply to all recipients of the current item.|
| **replyToFolder**|Post a reply to a folder.|

## Modifying a Built-in Action

After you have identified a built-in action in an  **action** element, you can modify it by specifying optional child elements and attribute for the **action** element.

 [To specify the message class and the form for the resulting item](specify-the-message-class-and-form-for-an-item-resulting-from-an-action.md)

 [To specify the way that the body of the current item is included in the resulting item](include-the-original-body-in-an-item-resulting-from-an-action.md)

 [To specify the way that Outlook responds when executing the action](specify-the-way-outlook-responds-when-executing-an-action.md)

 [To specify the prefix value in the subject of the resulting item](specify-a-subject-prefix-of-an-item-resulting-from-an-action.md)


### To disable the built-in action for the form region




1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the action as the value of the  **name** attribute of the **action** element.
    
3. Specify  **true** as the value of the disabled attribute of the **action** element.
    
The following example identifies the built-in action,  `replyToFolder`, and disables it:




```
<customActions>
    <action name="replyToFolder" disabled="true">
    </action>
</customActions>
```


 **Note**  You can assign  **disabled** either a string value or an integer value. The default value is **false** or **0**. To disable a built-in action for a form region, assign  **disabled** either **true** or **1**.


