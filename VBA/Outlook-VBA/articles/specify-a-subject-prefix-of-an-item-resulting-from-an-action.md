---
title: Specify a Subject Prefix of an Item Resulting from an Action
ms.prod: outlook
ms.assetid: a293f15e-ef68-84fe-2ef6-9badbfb9b194
ms.date: 06/08/2017
---


# Specify a Subject Prefix of an Item Resulting from an Action

## To specify the prefix value in the subject of the resulting item


1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the action as the value of the  **name** attribute of the **action** element.
    
3. Specify a string that represents the prefix of the subject line of the resulting item as the value of the child  **subjectPrefix** element of the **action** element.
    
The following example assigns  `replyToBlog` as the internal name of a custom action, and specifies `Re` as the subject line prefix for the resulting item:


```
<customActions>
    <action name="replyToBlog">
        <subjectPrefix>Re</subjectPrefix>
        <!-- Further characterize this action -->
    </action>
</customActions>
```


