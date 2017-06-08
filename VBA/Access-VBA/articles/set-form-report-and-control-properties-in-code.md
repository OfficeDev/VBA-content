---
title: Set Form, Report, and Control Properties in Code
ms.prod: access
ms.assetid: 23d88ab3-9ee6-5f7f-2351-14bb94d7a27b
ms.date: 06/08/2017
---


# Set Form, Report, and Control Properties in Code

 **[Form](form-object-access.md)**, **[Report](report-object-access.md)**, and **[Control](control-object-access.md)** objects are Access objects. You can set properties for these objects from within a **Sub**, **Function**, or event procedure. You can also set properties for form and report sections.


## To set a property of a form or report

Refer to the individual form or report within the  **[Forms](forms-object-access.md)** or **[Reports](reports-object-access.md)** collection, followed by the name of the property and its value. For example, to set the **[Visible](form-visible-property-access.md)** property of the Customers form to **True** (-1), use the following line of code:


```vb
Forms!Customers.Visible = True
```

You can also set a property of a form or report from within the object's module by using the object's  **Me** property. Code that uses the **Me** property executes faster than code that uses a fully qualified object name. For example, to set the **[RecordSource](form-recordsource-property-access.md)** property of the Customers form to an SQL statement that returns all records with a CompanyName field entry beginning with "A" from within the Customers form module, use the following line of code:




```vb
Me.RecordSource = "SELECT * FROM Customers " _ 
    &; "WHERE CompanyName Like 'A*'"
```


## To set a property of a control

Refer to the control in the  **[Controls](form-controls-property-access.md)** collection of the **Form** or **Report** object on which it resides. You can refer to the **Controls** collection either implicitly or explicitly, but the code executes faster if you use an implicit reference. The following examples set the **Visible** property of a text box called CustomerID on the Customers form:


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

Refer to the form or report within the  **Forms** or **Reports** collection, followed by the **Section** property and the integer or constant that identifies the section. The following examples set the **Visible** property of the page header section of the Customers form to **False**:


```vb
Forms!Customers.Section(3).Visible = False
```


```vb
Me!Section(acPageHeader).Visible = False
```


- For each property you want to set, you can look up the property in the Help index to find information about:
    
      - Whether you can set the property from Visual Basic.
    
  - Views in which you can set the property. Not all properties can be set in all views. For example, you can set a form's  **BorderStyle** property only in form Design view.
    
  - Which values you should use to set the property. You often use different settings when you set a property in Visual Basic instead of in the property sheet. For example, if the property settings are selections from a list, you must use the value or numeric equivalent for each selection.
    
- To set default properties for controls from Visual Basic, use the  **DefaultControl** property.
    

