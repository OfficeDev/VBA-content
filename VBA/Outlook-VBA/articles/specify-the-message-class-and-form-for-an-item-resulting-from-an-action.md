---
title: Specify the Message Class and Form for an Item Resulting from an Action
ms.prod: outlook
ms.assetid: 89cb6501-3d48-3bcb-c29d-e2e56274f6cb
ms.date: 06/08/2017
---


# Specify the Message Class and Form for an Item Resulting from an Action

## To specify the message class and the form for the resulting item


1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the action as the value of the  **name** attribute of the **action** element.
    
3. Specify a string that represents the message class of the resulting item as the value of the child  **targetForm** element of the **action** element.
    
The following example assigns  `replyToBlog` as the internal name of a custom action, and `IPM.Post` as the message class of the resulting item. The resulting item will use the same form that a contact item uses by default:


```XML
<customActions>
    <action name="replyToBlog">
        <targetForm>IPM.Post</targetForm>
        <!-- Further characterize this action -->
    </action>
</customActions>
```


 **Note**  You can specify  `this` as the value of the **targetForm** element to use the same message class and same form as those that the form region is defined on.


