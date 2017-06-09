---
title: Program with Class Modules
keywords: vbaac10.chm5187902
f1_keywords:
- vbaac10.chm5187902
ms.prod: access
ms.assetid: 6b10be38-bfe6-dea2-4aa5-4859722c1869
ms.date: 06/08/2017
---


# Program with Class Modules

In Access, there were two types of modules: standard modules and class modules. In Access 95, class modules existed only in association with a form or report. In Access 97, they also existed on the  **Modules** tab of the Database window.


## Creating Custom Objects with Class Modules

You can use a class module to create a definition for a custom object. The name with which you save the class module becomes the name of your custom object. Public  **Sub** and **Function** procedures that you define within a class module become custom methods of the object. Public **Property Let**, **Property Get**, and **Property Set** procedures become properties of the object.

Once you've defined procedures within the class module, you can create the new object by creating a new instance of the class. To create a new instance of a class, you declare a variable of the type defined by that class. For example, if the name of your class is ABasicClass, you would create a new instance of it in the following manner:




```vb
Dim abc As New ABasicClass
```

When you run the code containing this declaration, Visual Basic creates the new instance. You can then apply its methods and properties by using the variable . For example, if you've defined a custom method called ListNames, you could apply it as follows:




```
abc.ListNames
```


## New in Access 95: Creating the Default Instance of a Form Class

When you open a form in Form view, whether from the user interface or from Visual Basic, you create an instance of that form's class module. In other words, you designate space in memory where the object now exists, and you can then call its methods and set or return its properties from code, as you would for any built-in object. The same is true when you open a report in Print Preview.

When you refer to a form in Visual Basic code, you're usually working with the default instance of the form's class. A form's class has only one default instance. You can also create multiple instances of the same form's class from Visual Basic. When you create multiple instances of a form's class, you create nondefault instances.

There are four ways to create the default instance of a form. You can open an existing form by using the user interface, by executing the  **[OpenForm](docmd-openform-method-access.md)** method of the **[DoCmd](docmd-object-access.md)** object, by calling the **[CreateForm](application-createform-method-access.md)** method and switching the new form into Form view, or by using Visual Basic to create a variable of type **Form** to refer to the default instance. The following example opens an Employees form and points a **Form** object variable to it:




```vb
Dim frm As Form 
DoCmd.OpenForm "Employees" 
Set frm = Forms!Employees
```

Access also provides a shortcut that enables you to open a form and refer to a method or property of that form or one of its controls in one step. You refer to the form's class module as shown in the following example:




```vb
Form_Employees.Visible = True 
Form_Employees.Caption = "New Employees"
```

When you run this code, Access opens the Employees form in Form view if it's not already open and sets the form's caption to "New Employees." The form isn't visible until you explicitly set its  **[Visible](form-visible-property-access.md)** property to **True**. When the procedure that calls this code has finished executing, this instance of the form is destroyed; that is, the form is closed.

If you try to run this code when the Employees form is open in Design view, Access generates a run-time error. The form must either be open in Form view or not open at all.

If you use this syntax to change a property of the form or one of its controls, that change is lost when the instance of the form is destroyed. This holds true any time you change a property setting for a form in Form view. You must change the property in Design view and save the change with the form.


## Creating Multiple Nondefault Instances of Forms

You can create multiple nondefault instances of a form's class if you want to display more than one instance of a form at a time. For example, you might want to display the records for an employee and the employee's manager at the same time. You can create one instance of the Employees form's class to display the employee's record, and one to display the manager's record.

To create new, nondefault instances of a form's class from Visual Basic, declare a variable for which the type is the name of the form class module. You must include the  **New** keyword in the variable declaration. For example, the following code creates a new instance of the Employees form and assigns it to a variable of type **Form**:




```vb
Dim frm As New Form_Employees
```

This nondefault instance of the form isn't visible until you explicitly set its  **Visible** property.

When the procedure that creates this instance has finished executing, the instance is removed from memory unless you've declared the variable representing it as a module-level variable. Since module-level variables retain their values until they are reset with the  **Reset** command on the **Run** menu or the **Reset** button on the toolbar, the form will stay open if the variable has been declared as a module-level variable.

Any properties that you set will affect this instance of the form's class, but won't be saved with the form. Also, a new instance of the form's class can't be created if the form is open in Design view.

A nondefault instance of a form's class can't be referred to by name in the  **[Forms](forms-object-access.md)** collection. You can refer to it by index number only. Since you can create multiple nondefault instances of a form, and each instance has the same name, you can have more than one form with the same name in the **Forms** collection, without any means of distinguishing them other than by index number.


