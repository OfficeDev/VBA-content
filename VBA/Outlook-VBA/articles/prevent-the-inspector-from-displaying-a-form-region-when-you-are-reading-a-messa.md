---
title: Prevent the Inspector from Displaying a Form Region When You are Reading a Message
ms.prod: outlook
ms.assetid: f84c5797-c24f-4f16-4135-c4f1999c6aba
ms.date: 06/08/2017
---


# Prevent the Inspector from Displaying a Form Region When You are Reading a Message

When you create a form region in a custom form for mail or post items, by default, the form region will be displayed in the inspector when you read a mail or post item that uses that custom form. If you want to prevent the inspector from displaying the form region, specify this in the form region manifest XML file that you register for the form region. For more information on registering a form region, see  [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).


## To prevent the inspector from displaying a form region while in read mode


- In the form region manifest XML file, specify  **false** as the value of the **showInspectorRead** element.
    
The following example disables the inspector from displaying a form region when in read mode:


```
<showInspectorRead>false</showInspectorRead>
```


 **Note**  You can assign  **showInspectorRead** either a string value or an integer value. The default value is **true** or **1**. To prevent the inspector from displaying the form region in read mode, assign either  **false** or **0**.


