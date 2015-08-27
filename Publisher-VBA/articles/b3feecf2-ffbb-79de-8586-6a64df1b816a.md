
# WizardProperties Object (Publisher)

 **Last modified:** July 28, 2015

Represents the settings available in a publication design or in a Design Gallery object's wizard.

## Example

Use the  ** [Properties](9f9811b3-10ee-d429-c5a2-8223349525f2.md)** property with a **Wizard** object to return a **WizardProperties** collection. The following example reports on the publication design associated with the active publication, displaying its name and current settings.


```
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 MsgBox "Publication Design associated with " _ 
 &amp; "current publication: " .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Note**  Depending on the language version of Microsoft Publisher that you are using, you may receive an error when using the above code. If this occurs, you will need to build in error handlers to circumvent the errors. For more information, see  ** [Wizard Object](c0a64ee9-d1fa-6dc7-5221-ff2d32874ea0.md)**.

