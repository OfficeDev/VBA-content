
# ValidationRuleSet.Delete Method (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Deletes the  ** [ValidationRuleSet](cd2fc58a-5d7c-cf31-7aab-41bdeee9f105.md)** object from the document.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Delete**

 _expression_A variable that represents a  **ValidationRuleSet** object.


### Return Value

 **Nothing**


## Remarks
<a name="sectionSection1"> </a>

Calling the  **Delete** method also deletes all ** [ValidationRule](c9efb9b4-10b0-b6aa-cc78-2a01fd3e8357.md)** objects that are associated with the validation rule set.


## Example
<a name="sectionSection2"> </a>

The following sample code is based on code provided by:  [David Parker](http://www.bvisual.net)

The following Visual Basic for Applications (VBA) example shows how to use the  **Delete** method to delete a validation rule set named "Fault Tree Analysis" from the active document.




```
' Delete a rule set from the active document.
Public Sub Delete_Example()

    Dim strValidationRuleSetNameU As String
    strValidationRuleSetNameU = "Fault Tree Analysis"
    
    ActiveDocument.Validation.RuleSets(strValidationRuleSetNameU).Delete
   
End Sub
```

