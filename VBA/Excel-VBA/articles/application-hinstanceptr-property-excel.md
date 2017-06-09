---
title: Application.HinstancePtr Property (Excel)
keywords: vbaxl10.chm133334
f1_keywords:
- vbaxl10.chm133334
ms.prod: excel
api_name:
- Excel.Application.HinstancePtr
ms.assetid: fddc40e9-08fc-34ef-60b2-41e8afa86575
ms.date: 06/08/2017
---


# Application.HinstancePtr Property (Excel)

Returns a handle to the instance of Excel represented by the specified  **[Application](application-object-excel.md)** object. Read-only **Variant** .


## Syntax

 _expression_ . **HinstancePtr**

 _expression_ A variable that represents an **Application** object.


## Remarks

This property returns a correct handle in both the 32- and 64-bit versions of Excel. It extends the functionality of the  **[Hinstance](application-hinstance-property-excel.md)** property of the **Application** object, which only works correctly in the 32-bit version of Excel.

The ideal data type to use with this property is the  **[LongPtr](http://msdn.microsoft.com/library/10ee4c07-b686-5b86-5cea-250a9218e7ba%28Office.15%29.aspx)** data type. Assigning the value returned by this property to a **LongPtr** variable will work as expected in both 32- and 64-bit versions of Excel. The property is defined as **Variant** for internal implementation reasons. However, it always returns a 32-bit value on 32-bit systems and a 64-bit value on 64-bit systems.

This property only works starting with Excel, and is only required with the 64-bit version of Excel. If you must write code that will also work with earlier versions of Excel, in order to avoid compilation errors, read this property under an  `#if Win64` conditional compilation directive, and use the **Hinstance** property under the `#else` directive.

Note that this property works fine in both 32- and 64-bit environments starting with Excel. Therefore, if your code is intended to be used only with Excel or later, either 32- or 64-bit, it can read this property without conditional compilation.

For more information about how to use VBA in 64-bit environments, see [64-Bit Visual Basic for Applications Overview](http://msdn.microsoft.com/library/a44e016f-1019-300e-5150-916ff32f70c1%28Office.15%29.aspx).


## Example

In this example, a message box displays the 

Excel instance handle to the user.




```vb
Sub CheckHinstance() 
    MsgBox Application.HinstancePtr 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

