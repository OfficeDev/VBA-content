
# Application.WindowActivate Event (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Occurs when the application window or any document window is activated.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **WindowActivate**( **_Pres_**,  **_Wn_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Pres|Required| **Presentation**|The presentation displayed in the activated window.|
|Wn|Required| **DocumentWindow**|The activated document window.|

## Remarks
<a name="sectionSection1"> </a>

For information about using events with the  **Application** object, see [How to: Use Events with the Application Object](b657ab62-67fa-4eeb-736c-86e31a026c73.md).

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this event maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.EApplication_WindowActivateEventHandler** (the **WindowActivate** delegate.)
    
-  **Microsoft.Office.Interop.PowerPoint.EApplication_Event.WindowActivate** (the **WindowActivate** event.)
    

## Example
<a name="sectionSection2"> </a>

This example opens every activated presentation in slide sorter view.


```
Private Sub App_WindowActivate (ByVal Pres As Presentation, ByVal Wn As DocumentWindow) 
    Wn.ViewType = ppViewSlideSorter 
End Sub 

```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Application Object](978c2b99-4271-b953-4283-73b5f3d96f41.md)
#### Other resources


 [Application Object Members](7a9042da-ef77-ebba-c872-f736bf486674.md)
