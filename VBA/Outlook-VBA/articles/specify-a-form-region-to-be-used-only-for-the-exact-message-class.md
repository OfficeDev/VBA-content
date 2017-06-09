---
title: Specify a Form Region to be Used Only for the Exact Message Class
ms.prod: outlook
ms.assetid: cf08e1da-bc82-8f8f-0790-09bbf24bc8cd
ms.date: 06/08/2017
---


# Specify a Form Region to be Used Only for the Exact Message Class

When you register a form region for a message class, by default, the inspector displays the form region for any item belonging to that message class, as well as any item that belongs to a derived message class. For example, if you create a form region and register it for  **IPM.Appointment**, the inspector will display that form region for any item that belongs to  **IPM.Appointment**, and any item that belongs to a message class derived from  **IPM.Appointment**, such as  **IPM.Appointment.Customers**. If you want the inspector to use the form region for only the exact message class that the form region is registered for, specify this in the form region manifest XML file for the form region. For more information on registering a form region, see  [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).


## To allow a form region to be used for only the exact message class


- In the form region manifest XML file, specify  **true** as the value of the **exactMessageClass** element.
    
The following example disables a form region from being modified in the Forms Designer:


```
<exactMessageClass>true</exactMessageClass>
```


 **Note**  You can assign  **exactMessageClass** either a string value or an integer value. The default value is **false** or **0**. To allow a form region to be used for only the exact message class, assign either  **true** or **1**.


