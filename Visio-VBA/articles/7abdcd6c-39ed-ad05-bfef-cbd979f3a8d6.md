
# Document.FooterCenter Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Gets or sets the text string that appears in the center portion of a document's footer. Read/write.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **FooterCenter**

 _expression_A variable that represents a  **Document** object.


### Return Value

String


## Remarks
<a name="sectionSection1"> </a>

You can also set this value in the  **Center** box under **Footer** in the **Header and Footer** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, and then in the  **Preview** group, click **Header &amp; Footer**).

Both the string returned by the property and the string you pass to the property can contain escape codes that represent data. These escape codes can be concatenated with other text. For a list of valid escape codes you can use with the  **FooterCenter** property, see the ** [FooterLeft](e832c09d-3ddb-4351-43ad-e1c5633b7bc9.md)** property topic.


## Example
<a name="sectionSection2"> </a>

The following macro shows how to place a string containing the current page number and total number of pages into the center portion of a document's footer. After you run this macro on a one-page document, the center portion of the footer contains "Page 1 of 1".


```
 
Sub FooterCenter_Example()  
 
    Dim strFooter as String 
 
    'Build the footer string.  
    strFooter = "Page &amp;p of &amp;P"  
 
    'Set the footer of the current document.  
     ThisDocument.FooterCenter = strFooter 
  
End Sub 

```

