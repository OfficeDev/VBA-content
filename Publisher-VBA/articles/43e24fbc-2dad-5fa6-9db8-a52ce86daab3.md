
# ShapeRange.Wizard Property (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [Wizard](c0a64ee9-d1fa-6dc7-5221-ff2d32874ea0.md)**object representing the publication design associated with the specified publication or the wizard associated with the specified Design Gallery object.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Wizard**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks
<a name="sectionSection1"> </a>

When accessing the  **Wizard** property from the **Document** or **Page** object, if the specified publication is not associated with any publication design, an error occurs. When accessing the **Wizard** property from the **Shape** or **ShapeRange** object, if the specified object is not a Design Gallery object, an error occurs.


## Example
<a name="sectionSection2"> </a>

The following example reports on the publication design associated with the active publication, displaying its name and current settings.


```
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 Debug.Print "Publication design associated with " _ 
 &amp; "current publication: " _ 
 &amp; .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Setting: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Note**  Depending on the language version of Publisher that you are using, you may receive an error when using the above code. If this occurs, you will need to build in error handlers to circumvent the errors. For more information, see  ** [Wizard](c0a64ee9-d1fa-6dc7-5221-ff2d32874ea0.md)** object .

