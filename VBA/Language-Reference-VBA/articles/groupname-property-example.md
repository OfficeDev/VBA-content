---
title: GroupName Property Example
keywords: fm20.chm5225162
f1_keywords:
- fm20.chm5225162
ms.prod: office
ms.assetid: cff11547-2c4a-e8b6-294f-fc0fc2c06e88
ms.date: 06/08/2017
---


# GroupName Property Example

The following example uses the  **GroupName** property to create two groups of **OptionButton** controls on the same form.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains five  **OptionButton** controls named OptionButton1 through OptionButton5.



```vb
Private Sub UserForm_Initialize() 
 OptionButton1.GroupName = "Widgets" 
 OptionButton2.GroupName = "Widgets" 
 OptionButton4.GroupName = "Widgets" 
 
 OptionButton3.GroupName = "Gadgets-Group2" 
 OptionButton5.GroupName = "Gadgets-Group2" 
End Sub
```


