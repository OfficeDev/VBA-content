---
title: Create Multiple Instances of a Form
ms.prod: access
ms.assetid: 1e59bcce-f65f-8632-d96c-9e93af419d05
ms.date: 06/08/2017
---


# Create Multiple Instances of a Form

Sometimes it is useful to display more than one instance of a form at a time. For example, you might want to display the records for an employee and the employee's manager at the same time. You can create one instance of the Employees form's class to display the employee's record, and one to display the manager's record.

When you create a new instance of a form or report class, the new instance has all the properties and methods of a  **Form** or **Report** object, and its properties are set to the same values as those in the original **Form** or **Report** object. Additionally, any procedures that you have written in the form or report class module behave as methods and properties of the new instance.

To create a new instance of a form or report class, you declare a new object variable by using the  **[Shell](http://msdn.microsoft.com/library/033bffb0-540f-2c17-2aed-d25d10bedd8c%28Office.15%29.aspx)** keyword and the name of the form's or report's class module. The name of the class module appears in the title bar of the module. It indicates whether the class is associated with a form or a report and includes the name of the form or report. For example, the class name for an Employee form is Form_Employees. The following line of code creates a new instance of the Employees form:




```vb
Dim frmInstance As New Form_Employees 

```

By creating multiple instances of an Employees form class, you could show information about one employee on one form instance, and show information about another employee on another form instance. 

 **Note**  When you create an instance of a form class by using the  **New** keyword, it is hidden. To show the form, set the **[Visible](form-visible-property-access.md)** property to **True**.

You should declare the variable that represents the new instance of a form class at the module level. If you declare the variable at the procedure level, the variable goes out of scope when the procedure finishes running, and the new instance is removed from memory. The instance exists in memory only as long as the variable to which it is assigned remains in scope.
Any properties that you set will affect this instance of the form's class, but will not be saved with the form. Also, a new instance of the form's class cannot be created if the form is open in Design view.

