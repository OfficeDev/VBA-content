---
title: Prevent a Replacement Form Region from Being Used to Create a New Item or from Being Modified in the Forms Designer
ms.prod: outlook
ms.assetid: af7ea177-329f-1e96-287a-392a4780ff2a
ms.date: 06/08/2017
---


# Prevent a Replacement Form Region from Being Used to Create a New Item or from Being Modified in the Forms Designer

 Through the **Actions** menu or the **Choose Form** dialog, you can select a replacement or replace-all form region in the current folder to create a new item. Through the drop down list of the **Design Form** dialog, you can also select a replacement or replace-all form region and modify it in the Forms Designer. If you want to prevent a replacement or replace-all form region from being displayed in the above dialog boxes and menu, specify so in the form region manifest XML file that you register for the form region. For more information on registering a form region, see [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).


## To prevent a form region from being modified in the Forms Designer


- In the form region manifest XML file, specify  **false** as the value of the **hidden** element.
    
The following example disables a replacement form region from being used to create a new item or from being modified in the Forms Designer:


```
<hidden>false</hidden>
```


 **Note**  You can assign either a string value or an integer value to  **hidden**. The default value is  **false** or **0**. To prevent a form region from being used to create a new item or from being modified in the Forms Designer, assign either  **true** or **1**.


