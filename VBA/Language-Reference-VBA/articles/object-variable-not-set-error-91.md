---
title: Object variable not set (Error 91)
keywords: vblr6.chm1000091
f1_keywords:
- vblr6.chm1000091
ms.prod: office
ms.assetid: db8be8b0-9437-d53e-18b9-1d646b40ea66
ms.date: 06/08/2017
---


# Object variable not set (Error 91)

There are two steps to creating an [object variable](vbe-glossary.md). First you must declare the object variable. Then you must assign a valid reference to the object variable using the  **Set** statement. Similarly, a **With...End With** block must be initialized by executing the **With** statement entry point. This error has the following causes and solutions:


- You attempted to use an object variable that isn't yet referencing a valid object.
    
    Specify or respecify a reference for the object variable. For example, if the  **Set** statement is omitted in the following code, an error would be generated on the reference to `MyObject`:
    


```vb
Dim MyObject As Object    ' Create object variable. 
Set MyObject = Sheets(1)    ' Create valid object reference. 
MyCount = MyObject.Count    ' Assign Count value to MyCount. 

  ```

- You attempted to use an object variable that has been set to  **Nothing**.
    
```vb
Set MyObject = Nothing    ' Release the object. 
MyCount = MyObject.Count    ' Make a reference to a released object. 

  ```


    Respecify a reference for the object variable. For example, use a new  **Set** statement to set a new reference to the object.
    
- The object is a valid object, but it wasn't set because the [object library](vbe-glossary.md) in which it is described hasn't been selected in the **Add References** dialog box.
    
    Select the object library in the  **Add References** dialog box.
    
- The target of a  **GoTo** statement is inside a **With** block.
    
    Don't jump into a  **With** block. Make sure the block is initialized by executing the **With** statement entry point.
    
- You specified a line inside a  **With** block when you chose the **Set Next Statement** command.
    
    The  **With** block must be initialized by executing the **With** statement.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).


