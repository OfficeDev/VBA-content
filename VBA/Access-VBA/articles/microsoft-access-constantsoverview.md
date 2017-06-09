---
title: Microsoft Access Constants - Overview
keywords: vbaac10.chm4052
f1_keywords:
- vbaac10.chm4052
ms.prod: access
ms.assetid: 95a4bf62-7102-d2c4-5d66-f28e49c21707
ms.date: 06/08/2017
---


# Microsoft Access Constants - Overview

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[General](#sectionSection0)
[Symbolic Constants](#sectionSection1)
[Intrinsic Constants](#sectionSection2)
[System-Defined Constants](#sectionSection3)
[General](#sectionSection4)
[Symbolic Constants](#sectionSection5)
[Intrinsic Constants](#sectionSection6)
[System-Defined Constants](#sectionSection7)



## General
<a name="sectionSection0"> </a>

A constant represents a numeric or string value that doesn't change. You can use constants to improve the readability of your Visual Basic code and to make your code easier to maintain. In addition, the use of intrinsic constants ensures that code will continue to work even if the underlying values that the constants represent are changed in later releases of Microsoft Access.

Microsoft Access supports three types of constants:


- Symbolic constants, which you create by using the  **Const** statement and use in modules.
    
- Intrinsic constants, which are part of Microsoft Access or a referenced library.
    
- System-defined constants:  **True**, **False**, and **Null**.
    

## Symbolic Constants
<a name="sectionSection1"> </a>

Often, you'll use the same values repeatedly in your code, or you'll find that the code depends on certain numbers that have no obvious meaning. In these cases, you can make the code much easier to read and maintain by using symbolic or user-defined constants, which enable you to use a meaningful name in place of a number or string.

Once you have created a constant by using the  **Const** statement, you can't modify it or assign a new value to it. You also can't create a constant that has the same name as an intrinsic constant.

The following examples show some of the ways you can use the  **Const** statement to declare numeric and string constants:




```vb
Const conPI = 3.14159265                ' Pi equals this number. 
Const conPI2 = conPI * 2                ' A constant used to create another. 
Const conVersion = "Version 12.0"        ' Declare a string constant.
```


## Intrinsic Constants
<a name="sectionSection2"> </a>

In addition to the constants you declare with the  **Const** statement, Microsoft Access automatically declares a number of intrinsic constants and provides access to the Visual Basic for Applications (VBA) constants, and ActiveX Data Objects (ADO) constants. You can also use constants in other referenced object libraries. For more information on adding references see[Set References to Type Libraries](http://msdn.microsoft.com/library/6314a89b-89e9-d8c1-5964-889a361afcd1%28Office.15%29.aspx).

Any intrinsic constant can be used in a macro or Visual Basic. These constants are available at all times. The specific built-in constants used with a particular function, method, or property are described in the Help topic for that function, method, or property.


 **Note**  You can use the Object Browser to view lists of intrinsic constants from all available object libraries.

Intrinsic constants have a two letter prefix identifying the object library that defines the constant. Constants from the Microsoft Access library are prefaced with "ac"; constants from the ADO library are prefaced with "ad"; and constants from the Visual Basic library are prefaced with "vb". For example:


-  **acForm**
    
-  **adAddNew**
    
-  **vbCurrency**
    

 **Note**  Because the values represented by the intrinsic constants may change in future versions of Microsoft Access, you should use the constants instead of their actual values. You can, however, display the actual value of a constant by choosing the constant in the Object Browser or by typing ? _constantname_ in the Immediate window.

You can use intrinsic constants wherever you can use symbolic, or user-defined constants, including in expressions. The following example shows how you might use the intrinsic constant  **vbCurrency** to determine whether the variable is a **Variant** for which the **VarType** function returns 6 ( **Currency** ):




```vb
Dim varNum As Variant 
 
If VarType(varNum) = vbCurrency Then 
    Debug.Print "varNum contains Currency data." 
Else 
    Debug.Print "varNum doesn't contain Currency data." 
End If
```


## System-Defined Constants
<a name="sectionSection3"> </a>

You can use the system-defined constants  **True**, **False**, and **Null** anywhere in Microsoft Access. For example, you can use **True** in the following macro condition expression. The condition is met if the **Visible** property setting for the Employees form equals **True**.


```vb
Forms!Employees.Visible = True
```

You can use the constant  **Null** anywhere in Microsoft Access. For example, you can use **Null** to set the **DefaultValue** property for a form control by using the following expression:




```
=Null
```


## General
<a name="sectionSection4"> </a>

A constant represents a numeric or string value that doesn't change. You can use constants to improve the readability of your Visual Basic code and to make your code easier to maintain. In addition, the use of intrinsic constants ensures that code will continue to work even if the underlying values that the constants represent are changed in later releases of Microsoft Access.

Microsoft Access supports three types of constants:


- Symbolic constants, which you create by using the  **Const** statement and use in modules.
    
- Intrinsic constants, which are part of Microsoft Access or a referenced library.
    
- System-defined constants:  **True**, **False**, and **Null**.
    

## Symbolic Constants
<a name="sectionSection5"> </a>

Often, you'll use the same values repeatedly in your code, or you'll find that the code depends on certain numbers that have no obvious meaning. In these cases, you can make the code much easier to read and maintain by using symbolic or user-defined constants, which enable you to use a meaningful name in place of a number or string.

Once you have created a constant by using the  **Const** statement, you can't modify it or assign a new value to it. You also can't create a constant that has the same name as an intrinsic constant.

The following examples show some of the ways you can use the  **Const** statement to declare numeric and string constants:




```vb
Const conPI = 3.14159265                ' Pi equals this number. 
Const conPI2 = conPI * 2                ' A constant used to create another. 
Const conVersion = "Version 12.0"        ' Declare a string constant.
```


## Intrinsic Constants
<a name="sectionSection6"> </a>

In addition to the constants you declare with the  **Const** statement, Microsoft Access automatically declares a number of intrinsic constants and provides access to the Visual Basic for Applications (VBA) constants, and ActiveX Data Objects (ADO) constants. You can also use constants in other referenced object libraries. For more information on adding references see[Set References to Type Libraries](http://msdn.microsoft.com/library/6314a89b-89e9-d8c1-5964-889a361afcd1%28Office.15%29.aspx).

Any intrinsic constant can be used in a macro or Visual Basic. These constants are available at all times. The specific built-in constants used with a particular function, method, or property are described in the Help topic for that function, method, or property.


 **Note**  You can use the Object Browser to view lists of intrinsic constants from all available object libraries.

Intrinsic constants have a two letter prefix identifying the object library that defines the constant. Constants from the Microsoft Access library are prefaced with "ac"; constants from the ADO library are prefaced with "ad"; and constants from the Visual Basic library are prefaced with "vb". For example:


-  **acForm**
    
-  **adAddNew**
    
-  **vbCurrency**
    

 **Note**  Because the values represented by the intrinsic constants may change in future versions of Microsoft Access, you should use the constants instead of their actual values. You can, however, display the actual value of a constant by choosing the constant in the Object Browser or by typing ? _constantname_ in the Immediate window.

You can use intrinsic constants wherever you can use symbolic, or user-defined constants, including in expressions. The following example shows how you might use the intrinsic constant  **vbCurrency** to determine whether the variable is a **Variant** for which the **VarType** function returns 6 ( **Currency** ):




```vb
Dim varNum As Variant 
 
If VarType(varNum) = vbCurrency Then 
    Debug.Print "varNum contains Currency data." 
Else 
    Debug.Print "varNum doesn't contain Currency data." 
End If
```


## System-Defined Constants
<a name="sectionSection7"> </a>

You can use the system-defined constants  **True**, **False**, and **Null** anywhere in Microsoft Access. For example, you can use **True** in the following macro condition expression. The condition is met if the **Visible** property setting for the Employees form equals **True**.


```vb
Forms!Employees.Visible = True
```

You can use the constant  **Null** anywhere in Microsoft Access. For example, you can use **Null** to set the **DefaultValue** property for a form control by using the following expression:




```
=Null
```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

