---
title: "How to: Specify Possible Values for a List Box"
ms.prod: outlook
ms.assetid: 25ced223-0a3c-162a-127f-6b2f3ee9c5bc
ms.date: 06/08/2017
---


# How to: Specify Possible Values for a List Box

The following example fills a  **[ListBox](listbox-object-outlook-forms-script.md)** control with the values "Test1", "Test2", and "Test3" when you open the form.


```vb
Sub Item_Open() 
 
 ' Sets the name of page on the form, in this case, the 
 ' Message page on a MailItem form. 
 Set FormPage = Item.GetInspector.ModifiedFormPages("Message") 
 
 ' Sets Control to a list box called ListBox1. 
 Set Control = FormPage.Controls("ListBox1") 
 
 ' Assign values to the list box. 
 Control.PossibleValues = "Test1;Test2;Test3" 
 
End Sub
```


