
# How to: Change the Name of a Control

 **Last modified:** July 28, 2015

 _**Applies to:** Outlook 2013_

The following code example uses the  ** [ModifiedFormPages](ac377d47-846a-1217-592f-7ed190b824ca.md)** property of the current ** [Inspector](d7384756-669c-0549-1032-c3b864187994.md)** object to set the Microsoft Forms 2.0 **Name** property of a ** [CheckBox](1834855b-f96c-aaa1-24ce-81d1e4e4e1db.md)** on a page named "Test" to "Selection."




```
Item.GetInspector.ModifiedFormPages("Test").Checkbox1.Name = "Selection"
```


 **Note**  Each control should have a unique name.

