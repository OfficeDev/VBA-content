
# DoCmd.PrintOut Method (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


The  **PrintOut** method carries out the PrintOut action in Visual Basic.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **PrintOut**( **_PrintRange_**,  **_PageFrom_**,  **_PageTo_**,  **_PrintQuality_**,  **_Copies_**,  **_CollateCopies_**)

 _expression_A variable that represents a  **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PrintRange|Optional| **AcPrintRange**|A  ** [AcPrintRange](78d5a3d5-a94d-fb8c-45dd-5ba757576194.md)** constant that specifies the range to print. The default value is **acPrintAll**.|
|PageFrom|Optional| **Variant**|The first page to print. A numeric expression that's a valid page number in the active form or datasheet. This argument is required if you specify  **acPages** for theprintrange argument.|
|PageTo|Optional| **Variant**|The last page to print. A numeric expression that's a valid page number in the active form or datasheet. This argument is required if you specify  **acPages** for theprintrange argument.|
|PrintQuality|Optional| **AcPrintQuality**|A  ** [AcPrintQuality](5a4636c4-7034-34a8-3c75-7cd059b8f10a.md)** constant that specifies the print quality. the default value is **acHigh**.|
|Copies|Optional| **Variant**|The number of copies to print. If you leave this argument blank, the default (1) is assumed.|
|CollateCopies|Optional| **Variant**|Use  **True** (-1) to collate copies and **False** (0) to print without collating. If you leave this argument blank, the default ( **True**) is assumed.|

## Remarks
<a name="sectionSection1"> </a>

You can use the PrintOut action to print the active object in the open database. You can print datasheets, reports, forms, data access pages, and modules.


## Example
<a name="sectionSection2"> </a>

The following example prints two collated copies of the first four pages of the active form or datasheet:


```
DoCmd.PrintOut acPages, 1, 4, , 2
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [DoCmd Object](3ce44cca-9979-0a1e-9787-079a52ce528f.md)
#### Other resources


 [DoCmd Object Members](3e7ade9e-86e4-0751-188b-5d31c9101651.md)
