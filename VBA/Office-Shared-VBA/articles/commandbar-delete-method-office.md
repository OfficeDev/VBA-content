---
title: CommandBar.Delete Method (Office)
keywords: vbaof11.chm3004
f1_keywords:
- vbaof11.chm3004
ms.prod: office
api_name:
- Office.CommandBar.Delete
ms.assetid: 6976f273-dbd4-5f3d-52ef-0d6d5cc886c9
ms.date: 06/08/2017
---


# CommandBar.Delete Method (Office)

Deletes the  **CommandBar** object from the collection.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Delete**

 _expression_ Required. A variable that represents a **[CommandBar](commandbar-object-office.md)** object.


## Remarks

For the  **Scripts** collection, using the **Delete** method removes all scripts from the specified Microsoft Word document, Microsoft Excel worksheet, or Microsoft PowerPoint slide. A script anchor is represented by a shape in the host application. Therefore, the **Shape** object associated with each script anchor of type **msoScriptAnchor** is deleted from the **Shapes** collection in Excel and PowerPoint and from the **InlineShapes** and **Shapes** collections in Word.


## Example

This example deletes all custom command bars that are not visible.


```
foundFlag = False  
delBars = 0 
For Each bar In CommandBars 
    If (bar.BuiltIn = False) And _ 
    (bar.Visible = False) Then 
        bar.Delete 
        foundFlag =   
        delBars = delBars + 1 
    End If 
Next bar 
If Not foundFlag Then 
    MsgBox "No command bars have been deleted." 
Else 
    MsgBox delBars &amp; " custom bar(s) deleted." 
End If
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

