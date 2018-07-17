---
title: Form Name and Message Class Overview
ms.prod: outlook
ms.assetid: 8f72a998-b0c8-86ba-072b-5326ea785578
ms.date: 06/08/2017
---


# Form Name and Message Class Overview

The message class is an internal identifier that Microsoft Outlook and Microsoft Exchange use to locate and activate a form.

The message class property determines which folder to route the message to, and which form to activate to view the message. (The message class property corresponds to the MAPI property  [PidTagMessageClass](http://msdn.microsoft.com/library/1e704023-1992-4b43-857e-0a7da7bc8e87%28Office.15%29.aspx).) 

## Creating a message class when customizing form regions

Form regions give you more control over how Outlook handles your customized form. Depending on the type of form region that you use when you customize a form, you must create a unique message class for your form. For more information, see  [How to: Create a Form Region](create-a-form-region.md).


## Creating a message class when customizing form pages

First, some preliminaries. In the  **Publish Form As** dialog box, when you type a name in the **Display name** field, note that the **Form name** field reflects the display name by default. You can leave the form name the same as the display name or you can change the form name. The display name becomes the caption at the top of your form. The display name is also used to construct the name under which your form is published. When you publish your form, the display name is listed in the **Choose Form** dialog box.

Outlook automatically constructs a message class for the form that uses the name of the form, preceded by "IPM". For example, if you want to publish a mail message form named "MyForm", type "This is my Form" in the  **Display name** field, and type "MyForm" in the **Form name** field. At the bottom of the dialog box, Outlook displays the message class for your new form as: "IPM.Note.MyForm".

When you search in the  **Choose Form** dialog box, locate "This is my Form" displayed in the list. If you select it, the **Display name** field at the bottom of the dialog box displays, "This is my Form" and the **Form name** field displays, "MyForm".

Outlook automatically generates a message class from the form name and assigns it to the form. Then, when a form with that message class is selected, Outlook loads and displays an instance of that form. For example, Outlook uses the message class, "IPM.Note.MyForm" to locate the form with the display name, "This is my Form".


