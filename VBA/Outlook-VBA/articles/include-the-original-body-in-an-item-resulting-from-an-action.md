---
title: Include the Original Body in an Item Resulting from an Action
ms.prod: outlook
ms.assetid: 02806758-f126-2afd-2037-2a7a7292fb9d
ms.date: 06/08/2017
---


# Include the Original Body in an Item Resulting from an Action

## To specify the way that the body of the current item is included in the resulting item


1. In the form region manifest XML file, specify the action as a child  **action** element of the **customActions** element for that form region.
    
2. Specify the internal name of the action as the value of the  **name** attribute of the **action** element.
    
3. Specify a value for the child  **body** element of the **action** element.
    
The following example assigns  `replyToBlog` as the internal name of a custom action, and specifies that the body of the current item will be included and indented in the resulting item:


```
<customActions>
    <action name="replyToBlog">
        <body>indent</body>
        <!-- Further characterize this action -->
    </action>
</customActions>

```

Note that the child  **body** element of the **action** element can contain one of the folowing values:



| **Value**| **Description**|
| **attach**|The current item is attached to the resulting item.|
| **include**|The body of the current item is included as the body of the resulting item.|
| **indent**|The body of the current item is included in the body of the resulting item and indented.|
| **link**|A link to the current item is provided in the resulting item.|
| **omit**|The body of the current item is omitted from the resulting item.|
| **prefix**|The body of the current item is included in the body of the resulting item and prefixed with the quotation character.|
| **user**|The user's preferences are applied in how the body should be handled.|

