---
title: Specify the Way Outlook Responds When Executing an Action
ms.prod: outlook
ms.assetid: a2ea8dc3-728c-141b-42af-9b0a3c764a4a
ms.date: 06/08/2017
---


# Specify the Way Outlook Responds When Executing an Action

## To specify the way that Outlook responds when executing an action


1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the action as the value of the  **name** attribute of the **action** element.
    
3. Specify a value for the child  **method** element of the **action** element.
    
The following example assigns  `replyToBlog` as the internal name of a custom action, and specifies that Outlook will prompt the user to determine if he or she wants to open the resulting item now or send the item immediately:


```XML
<customActions>
    <action name="replyToBlog">
        <method>prompt</method>
        <!-- Further characterize this action -->
    </action>
</customActions>
```

Note that the child  **method** element of the **action** element can contain one of the following values:



| **Value**| **Description**|
| **open**|Outlook will open the resulting item in the inspector for the user to edit.|
| **prompt**|Outlook will prompt the user whether to open the resulting item now or send it immediately.|
| **send**|Outlook will send the resulting item automatically.|

