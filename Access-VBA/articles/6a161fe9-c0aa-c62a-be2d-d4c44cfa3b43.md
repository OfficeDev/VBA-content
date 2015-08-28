
# Call Procedures in a Subform or Subreport

 **Last modified:** July 28, 2015

 _**Applies to:** Access 2013_

You can call a procedure in a module associated with a subform or subreport in one of two ways. If the form containing the subform is open in Form view, you can refer to the procedure as a method on the subform. The following example shows how to call the procedure GetProductID in the Orders Subform, which is bound to a subform control on the Orders form:

In the Orders Subform class module enter:



```
Public Function GetProductID() As Variant 
 ' Return current productID. 
 GetProductID = ProductID 
End Function 
```




```
Forms!Orders![Orders Subform].Form.GetProductID
```

You can also create a new instance of the form that is being used as a subform, even if neither the main form nor the subform is open, and call the procedure. This will work for any form, whether or not it is being used as a subform. The following example shows how to create an instance of the Orders Subform and call the GetProductID procedure:



```
Dim frm As New [Form_Orders Subform] 
frm.GetProductID
```


 **Note**  When you create a new instance of a form with a name consisting of more than one word, enclose the class name of the form in brackets, as shown in the preceding example.

