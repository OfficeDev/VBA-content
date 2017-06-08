---
title: Displaying a Custom Dialog Box (Word)
keywords: vbawd10.chm5210530
f1_keywords:
- vbawd10.chm5210530
ms.prod: word
ms.assetid: edda05bb-092c-1352-671a-1349b58d5ba4
ms.date: 06/08/2017
---


# Displaying a Custom Dialog Box (Word)

To test your dialog box in the Visual Basic Editor, click  **Run Sub/UserForm** on the **Run** menu.

To display a dialog box from Visual Basic, use the  **Show**method. The following example displays the dialog box named UserForm1.



```vb
Private Sub GetUserName() 
 UserForm1.Show 
End Sub
```


 **Note**  Use the  **Unload** method in an event procedure, such as the Click event procedure for a command button, to close a dialog box.


