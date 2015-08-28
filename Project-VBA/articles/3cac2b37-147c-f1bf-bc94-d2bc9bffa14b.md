
# Application.GanttBarStyleDelete Method (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Deletes a Gantt bar style from the active Gantt Chart.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **GanttBarStyleDelete**( **_Item_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Item|Required| **String**| **String**. The name or row number of the Gantt bar to delete from the  **Bar Styles** dialog box.|

### Return Value

 **Boolean**


## Remarks
<a name="sectionSection1"> </a>

To manually show the  **Bar Styles** dialog box, click the **Format** tab under the **Gantt Chart Tools** tab. In the **Bar Styles** group, click **Bar Styles** in the **Format** drop-down list. The **Bar Styles** dialog box can contain up to 200 style entries.


## Example
<a name="sectionSection2"> </a>

The following command deletes style number 41 in the  **Bar Styles** dialog box.


```
GanttBarStyleDelete Item:="41"
```

