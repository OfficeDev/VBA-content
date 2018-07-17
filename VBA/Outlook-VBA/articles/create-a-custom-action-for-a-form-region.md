---
title: Create a Custom Action for a Form Region
ms.prod: outlook
ms.assetid: bf889270-3e80-a240-15e4-c57a3f1e7b9b
ms.date: 06/08/2017
---


# Create a Custom Action for a Form Region

You can specify custom actions for a form region. By default, four built-in actions exist for any form region. To create an action that suit your needs, you can either modify a built-in action or create a new custom action. For more information on modifying a built-in action, see  [How to: Modify a Built-in Action for a Form Region](modify-a-built-in-action-for-a-form-region.md).

When you create a custom action, you can specify the following characteristics in the form region manifest XML file that you register for the form region:

- The display name for the custom action.
    
- Whether the custom action will be displayed on the ribbon of an inspector.
    
- The way that the resulting item will be addressed.
    
- The message class of the item that results from executing the action.
    
- The way that the body of the current item is included in the resulting item.
    
- The way that Outlook responds when executing the action.
    
- The prefix value in the subject of the resulting item.
    
For more information on registering a form region, see  [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).

## Identifying Actions for a Form Region

Whether you choose to modify a built-in action or create a custom action, you define the action in the form region manifest XML file. Define these actions under the  **customActions** element, enclosing each action in its own **action** element and identifying it by the mandatory **name** attribute. The value of the **name** attribute is a string that represents the internal name of the action.


### To specify an internal name for an action


1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the action as the value of the  **name** attribute of the **action** element.
    
The following example assigns  `replyToBlog` as the internal name of one custom action, and `postToBlog` as the internal name of another custom action:


```
<customActions>
    <action name="replyToBlog">
        <!-- further characterize this action -->
    </action>
    <action name="postToBlog">
        <!-- further characterize this action -->
    </action>
</customActions>
```


## Defining a Custom Action

After you have identified an action in an  **action** element, you can further define the action by specifying optional child elements for the **action** element.


### To define a display name for the action




1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the action as the value of the  **name** attribute of the **action** element.
    
3. Specify the display name of the action as the value of the child  **title** element of the **action** element.
    
The following example assigns  `replyToBlog` as the internal name of a custom action, and `Reply to Blog` as the display name of the action:




```
<customActions>
    <action name="replyToBlog">
        <title>Reply to Blog</title>
        <!-- Further characterize this action -->
    </action>
</customActions>
```


### To specify that an action is to be displayed on the ribbon of an inspector




1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the action as the value of the  **name** attribute of the **action** element.
    
3. Specify  **true** as the value of the child **showOnRibbon** element of the **action** element.
    
The following example assigns  `replyToBlog` as the internal name of a custom action and specifies that it should not be displayed in the ribbon of an inspector:




```
<customActions>
    <action name="replyToBlog">
        <showOnRibbon>false</showOnRibbon>
        <!-- Further characterize this action -->
    </action>
</customActions>
```


 **Note**  You can assign  **showOnRibbon** either a string value or an integer value. Specifying **true** or **1** will display the action on the ribbon; specifying **false** or **0** will prevent it from being displayed on the ribbon.


### To specify the way that a resulting item will be addressed




1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the action as the value of the  **name** attribute of the **action** element.
    
3. Specify a value for the child  **addressLike** element of the **action** element.
    
The following example assigns  `replyToBlog` as the internal name of a custom action and specifies that the resulting new item will be addressed as a reply-all item, with all the original recipients copied over to the new item:




```
<customActions>
    <action name="replyToBlog">
        <addressLike>replyAll</addressLike>
        <!-- Further characterize this action -->
    </action>
</customActions>

```


 **Note**  The child  **addressLike** element of the **action** element can contain one of the following values:



| **Value**| **Description**|
| **forward**|Addresses the resulting item like a forward message that has no recipients specified. This also preserves attachments in the current item.|
| **reply**|Addresses the resulting item as a standard reply, with the sender specified in the  **To** line, and no one in the **CC** or **BCC** lines.|
| **replyAll**|Addresses the resulting item like a reply-all message, with all of the original recipients copied over to the resulting item.|
| **replyToFolder**|Addresses the resulting item like a post message to the current folder. This also clears the subject of the resulting item.|
| **response**|Addresses the resulting item as a response to vote, with the sender specified in the  **To** line, and no one in the **CC** or **BCC** lines.|

### 

 [To specify the message class and the form for the resulting item](specify-the-message-class-and-form-for-an-item-resulting-from-an-action.md)

 [To specify the way that the body of the current item is included in the resulting item](include-the-original-body-in-an-item-resulting-from-an-action.md)

 [To specify the way that Outlook responds when executing the action](specify-the-way-outlook-responds-when-executing-an-action.md)

 [To specify the prefix value in the subject of the resulting item](specify-a-subject-prefix-of-an-item-resulting-from-an-action.md)


