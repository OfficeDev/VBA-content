
# Application.VisualReportsSaveDatabase Method (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Saves a Visual Reports database to the default directory or to a specified directory.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **VisualReportsSaveDatabase**( **_strNamePath_**,  **_PjVisualReportsDataLevel_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|strNamePath|Optional| **String**|Name and full path of the location to which to save the database file (.mbd).|
|PjVisualReportsDataLevel|Optional| **Long**|Save data level. Can be one of the  ** [PjVisualReportsDataLevel](56792ea8-6459-38ef-e994-95024e6d8fe9.md)** constants. Default is **pjLevelAutomatic**.|

### Return Value

 **Boolean**


## Remarks
<a name="sectionSection1"> </a>

The PjVisualReportsDataLevel parameter specifies the level to which the timephased data can be accessed. For example, if  **pjLevelMonths** (months) is specified, it not possible to access **pjLevelDays** (days).


## Example
<a name="sectionSection2"> </a>

Following is an example of using The  **VisualReportsSaveDatabase** method.


```
Sub a() 
 Dim tf As Boolean 
 tf = Application.VisualReportsSaveDatabase("C:\mydb.mdb", pjLevelAutomatic) 
 If tf = True Then 
 MsgBox ("Database saved successfully") 
 Else 
 MsgBox ("Database wasn't saved successfully") 
 End If 
End Sub 

```

