
# InvisibleApp.AutoRecoverInterval Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Version Information](#sectionSection0)
 [Syntax](#sectionSection1)
 [Remarks](#sectionSection2)
 [Example](#sectionSection3)


Represents the time interval (in minutes) for how often you want to save copies of open documents that have unsaved changes in case of a power failure or an application error. Read/write.

## Version Information
<a name="sectionSection0"> </a>

Version Added: Visio 2000 SR-1 




## Syntax
<a name="sectionSection1"> </a>

 _expression_. **AutoRecoverInterval**

 _expression_A variable that represents an  **InvisibleApp** object.


### Return Value

Integer


## Remarks
<a name="sectionSection2"> </a>

Must be an integer value from zero (0) to 120, representing the interval in minutes. The default is 0. If the value of the  **AutoRecoverInterval** property is less than or equal to 0, no automatic recovery copies are created.

If the value of the  **AutoRecoverInterval** property is greater than 0, automatic recovery is enabled for all documents in the Microsoft Visio instance. To disable automatic recovery for a particular document, set its **AutoRecover** property to **False**.


## Example
<a name="sectionSection3"> </a>

The following Microsoft Visual Basic for Applications (VBA) macros show how to set the  **AutoRecoverInterval** property and how to use it to disable automatic recovery.


```
 
Public Sub AutoRecoverInterval_Example() 
  
    'Save automatic recovery copies of unsaved files 
    'every 10 minutes.  
    Application.AutoRecoverInterval = 10  
 
End Sub   
 
Public Sub DisableAutoRecover_Example() 
  
    'Tell Visio not to save automatic recovery copies of unsaved files.  
    Application.AutoRecoverInterval = 0  
 
End Sub 

```

