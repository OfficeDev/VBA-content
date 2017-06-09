---
title: Set Form, Report, and Control Properties in Visual Basic
keywords: vbaac10.chm5188061
f1_keywords:
- vbaac10.chm5188061
ms.prod: access
ms.assetid: 1f5b5f6b-b424-f35e-4add-21c45b5d74c4
ms.date: 06/08/2017
---


# Set Form, Report, and Control Properties in Visual Basic

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[To set a property of a form or report](#sectionSection0)
[To set a property of a control](#sectionSection1)
[To set a property of a form or report section](#sectionSection2)


 **[Form](http://msdn.microsoft.com/library/72EF9219-142B-B690-B696-3EBA9A5D4522%28Office.15%29.aspx)**, **[Report](http://msdn.microsoft.com/library/6F77C1B4-A9CE-7CAA-204C-FE0755C6F9DF%28Office.15%29.aspx)**, and **[Control](http://msdn.microsoft.com/library/CE2362E5-4390-590E-06C0-6F27E8D988CD%28Office.15%29.aspx)** objects are Microsoft Access objects. You can set properties for these objects from within a **Sub**, **Function**, or event procedure. You can also set properties for form and report sections.

## To set a property of a form or report
<a name="sectionSection0"> </a>

Refer to the individual form or report within the  **[Forms](http://msdn.microsoft.com/library/A41AF7BE-873C-EF8B-20CD-24B78A25B5CA%28Office.15%29.aspx)** or **[Reports](http://msdn.microsoft.com/library/37C5F55E-3C3A-6140-D305-7E8118D9D2B1%28Office.15%29.aspx)** collection, followed by the name of the property and its value. For example, to set the **Visible** property of the Customers form to **True** (-1), use the following line of code:


```vb
Forms!Customers.Visible = True
```

You can also set a property of a form or report from within the object's module by using the object's  **Me** property. Code that uses the **Me** property executes faster than code that uses a fully qualified object name. For example, to set the **RecordSource** property of the Customers form to an SQL statement that returns all records with a CompanyName field entry beginning with "A" from within the Customers form module, use the following line of code:




```
Me.RecordSource = "SELECT * FROM Customers " _ 
    &; "WHERE CompanyName Like 'A*'"
```


## To set a property of a control
<a name="sectionSection1"> </a>

Refer to the control in the  **[Controls](http://msdn.microsoft.com/library/26771888-86E8-28C3-6668-F793474CBB5B%28Office.15%29.aspx)** collection of the **Form** or **Report** object on which it resides. You can refer to the **Controls** collection either implicitly or explicitly, but the code executes faster if you use an implicit reference. The following examples set the **Visible** property of a text box called CustomerID on the Customers form:


```vb
' Faster method. 
Me!CustomerID.Visible = True
```


```vb
' Slower method. 
Forms!Customers.Controls!CustomerID.Visible = True
```

The fastest way to set a property of a control is from within an object's module by using the object's  **Me** property. For example, you can use the following code to toggle the **Visible** property of a text box called CustomerID on the Customers form:




```vb
With Me!CustomerID 
    .Visible = Not .Visible 
End With
```


## To set a property of a form or report section
<a name="sectionSection2"> </a>

Refer to the form or report within the  **Forms** or **Reports** collection, followed by the **Section** property and the integer or constant that identifies the section. The following examples set the **Visible** property of the page header section of the Customers form to **False**:


```vb
Forms!Customers.Section(3).Visible = False
```


```vb
Me!Section(acPageHeader).Visible = False
```


 **Note**  

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

