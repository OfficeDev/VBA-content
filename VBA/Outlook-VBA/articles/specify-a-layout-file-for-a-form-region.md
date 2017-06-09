---
title: Specify a Layout File for a Form Region
ms.prod: outlook
ms.assetid: fa418f65-a5e5-63fd-6efe-366268994711
ms.date: 06/08/2017
---


# Specify a Layout File for a Form Region

You can design and deploy a form region with or without an add-in. 

When you design a form region without an add-in, you create the form region in the Forms Designer and save the form region in an Outlook Form Storage file (.OFS). For more information on creating a form region, see  [How to: Create a Form Region](create-a-form-region.md).

To run the form region, you must register it in the Windows registry and specify the message class and the corresponding form region manifest XML file.


## To specify a file as the layout file for a form region


- In the form region manifest XML file, specify the file name of the .OFS file as the value of the  **layoutFile** element.
    
If you do not specify a full path name for the .OFS file, then Outlook will look for the file in the same folder as the form region manifest XML file that you have specified in the Windows registry for the form region. Otherwise, you can use environment variables in the  **layoutFile** element, such as the following:


```
<layoutFile>%ProgramFiles%\Addin\Addin.ofs</layoutFile>
```

You cannot use file paths expressed in the Universal Naming Convention (UNC) for the  **layoutFile** element.

If you use an add-in to design and deploy a form region, then you must specify the  **addin** element and must not specify the **layout** element.


