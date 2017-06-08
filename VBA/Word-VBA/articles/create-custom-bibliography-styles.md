---
title: Create Custom Bibliography Styles
ms.prod: word
ms.assetid: 4b55522f-3387-4e53-a8cd-3b616e3194c8
ms.date: 06/08/2017
---


# Create Custom Bibliography Styles

Create a custom bibliography style in Word by learning the steps (and XML code) you need to construct a simple custom style. Also, learn to make a more complex style file. Before we start, there is some information that you need to know:

The bibliography sources you create are all listed in the following file: \Microsoft\Bibliography\Sources.xml.

 **Note**  The \Bibliography\Sources.xml file won't exist until you create your first bibliography source in Word. All bibliography styles are stored in \Microsoft\Bibliography\Style.

## Building a basic bibliography style
<a name="Biblio_BuildBasicStyle"> </a>

First, create a basic bibliography style that the custom style will follow.


### Set up the bibliography style

To create a bibliography style, we will create an XML style sheet; that is, an .xsl file called MyBookStyle.xsl, using your favorite XML editor. Notepad will do fine. As the name suggests, our example is going to be a style for a "book" source type.

At the top of the file, add the following code:




```XML
<?xml version="1.0" ?> 

<!--List of the external resources that we are referencing-->
 
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
 
<!--When the bibliography or citation is in your document, it's just HTML-->
 
<xsl:output method="html" encoding="us-ascii"/>
   
<!--Match the root element, and dispatch to its children-->
   
<xsl:template match="/">

<xsl:apply-templates select="*" />

</xsl:template>
```

As the comments indicate, Word uses HTML to represent a bibliography or citation within a document. Most of the preceding XML code is just preparation for the more interesting parts of the style. For example, you can give your style a version number to track the changes you make, as shown in the following example.




```XML
<!--Set an optional version number for this style--> 

<xsl:template match="b:version"> 

   <xsl:text>2006.5.07</xsl:text>

</xsl:template>

```

More importantly, you can give your style a name. Add this tag: <xsl:when test="b:StyleNameLocalized">; and then give your style a name, in the language of your choice, by using the following code.




```XML
<xsl:when test="b:StyleNameLocalized/b:Lcid='1033'">

   <xsl:text>[Your Style Name]</xsl:text>
 
</xsl:when>
```

This section contains the locale name of your style. In the case of our example file, we want our custom bibliography style name, "Simple Book Style," to appear in the  **Style** drop-down list on the **References** tab. To do so, add the following XML code to specify that the style name be in the English locale (Lcid determines the language).




```XML
<!--Defines the name of the style in the References dropdown list-->
<xsl:when test="b:StyleNameLocalized"> 
   <xsl:choose> 
      <xsl:when test="b:StyleNameLocalized/b:Lcid='1033'"> 
         <xsl:text>Simple Book Style</xsl:text> 
      </xsl:when> 
</xsl:when>
```

Your style will now appear under its own name in the  **Bibliography Style** dropdown list-box in the application.

Now, examine the style details. Each source type in Word (for example, book, film, article in a periodical, and so forth) has a built-in list of fields that you can use for the bibliography. To see all the fields available for a given source type, on the  **References** tab, choose **Manage Sources**, and then in the  **Source Manager** dialog box, choose **New** to open the **Create Source** dialog box. Then select **Show All Bibliography Fields**.

A book source type has the following fields available:


- Author
    
- Title
    
- Year
    
- City
    
- State/Province
    
- Country/Region
    
- Publisher
    
- Editor
    
- Volume
    
- Number of Volumes
    
- Translator
    
- Short Title
    
- Standard Number
    
- Pages
    
- Edition
    
- Comments
    
In the code, you can specify the fields that are important for your bibliography style. Even when  **Show All Bibliography Fields** is cleared, these fields will appear and have a red asterisk next to them. For our book example, I want to ensure that the author, title, year, city, and publisher are entered, so I want a red asterisk to appear next to these fields to alert the user that these are recommended fields that should be filled out.




```XML
<!--Specifies which fields should appear in the Create Source dialog box when in a collapsed state (The Show All Bibliography Fields check box is cleared)-->

<xsl:template match="b:GetImportantFields[b:SourceType = 'Book']"> 
   <b:ImportantFields> 
      <b:ImportantField> 
         <xsl:text>b:Author/b:Author/b:NameList</xsl:text> 
      </b:ImportantField> 
      <b:ImportantField> 
         <xsl:text>b:Title</xsl:text> 
      </b:ImportantField> 
     <b:ImportantField> 
         <xsl:text>b:Year</xsl:text> 
      </b:ImportantField> 
      <b:ImportantField> 
         <xsl:text>b:City</xsl:text>
      </b:ImportantField> 
      <b:ImportantField> 
         <xsl:text>b:Publisher</xsl:text> 
      </b:ImportantField> 
   </b:ImportantFields> 
</xsl:template>
```

The text in the <xsl:text> tags are references to the Sources.xml file. These references pull out the data that will populate each of the fields. Examine Sources.xml in \Microsoft\Bibliography\Sources.xml) to get a better idea about how these references match up to what is in the XML file.


### Design the layout
<a name="Biblio_DesignLayout"> </a>

Output for bibliographies and citations is represented in a Word document as HTML, so to define how our custom bibliography and citation styles should look in Word, we'll have to add some HTML to our style sheet.

Suppose you want to format each entry in your bibliography in this manner:

**Last Name, First Name. (Year). Title. City: Publisher**

The HTML required to do this would be embedded in your style sheet as follows.




```XML
<!--Defines the output format for a simple Book (in the Bibliography) with important fields defined-->

<xsl:template match="b:Source[b:SourceType = 'Book']"> 

<!--Label the paragraph as an Office Bibliography paragraph-->

   <p> 
      <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/> 
      <xsl:text>, </xsl:text> 
      <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:First"/> 
      <xsl:text>. (</xsl:text> 
      <xsl:value-of select="b:Year"/> 
      <xsl:text>). </xsl:text> 
      <i> 
         <xsl:value-of select="b:Title"/> 
         <xsl:text>. </xsl:text> 
      </i> 
      <xsl:value-of select="b:City"/> 
      <xsl:text>: </xsl:text> 
      <xsl:value-of select="b:Publisher"/> 
      <xsl:text>.</xsl:text> 
   </p> 
</xsl:template>
```

When you reference a book source in your Word document, Word needs to access this HTML so that it can use the custom style to display the source, so you'll have to add code to your custom style sheet to enable Word to do this.




```XML
<!--Defines the output of the entire Bibliography-->
 
<xsl:template match="b:Bibliography"> 

   <html xmlns="http://www.w3.org/TR/REC-html40"> 
   
      <body> 

         <xsl:apply-templates select ="b:Source[b:SourceType = 'Book']"> 

         </xsl:apply-templates> 

      </body> 
   
   </html> 
</xsl:template>
```

In a similar fashion, you'll need to do the same thing for the citation output. Follow the pattern (Author, Year) for a single citation in the document.




```XML
<!--Defines the output of the Citation-->
<xsl:template match="b:Citation/b:Source[b:SourceType = 'Book']"> 
   <html xmlns="http://www.w3.org/TR/REC-html40"> 
      <body> 
         <!-- Defines the output format as (Author, Year)--> 
         <xsl:text>(</xsl:text> 
            <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/> 
         <xsl:text>, </xsl:text> 
         <xsl:value-of select="b:Year"/> 
         <xsl:text>)</xsl:text> 
      </body> 
   </html> 
</xsl:template>
```

Close up the file with the following lines.




```
<xsl:template match="text()" /> </xsl:stylesheet>
```

Save the file as MyBookStyle.XSL and drop it into the Styles directory (\Microsoft\Bibliography\Style). Restart Word, and your style is now under the style dropdown list. You can start using your new style.


## Create a complex style
<a name="Biblio_CreateComplexStyle"> </a>

One of the issues that complicate bibliography styles is that they often need to have a significant amount of conditional logic. For example, if the date is specified, you need to show the date, whereas if the date is not specified, you may need to use an abbreviation to indicate that there is no date for that source.

For a more specific example, in the APA style, if a date is not specified for a website source, the abbreviation "n.d." is used to denote no date, and the style should do this automatically. Here's an example:

APA website source with no date entered: Kwan, Y. (n.d.). Retrieved from www.microsoft.com APA website source with date entered: Kwan, Y. (2006, Jan 18). Retrieved from www.microsoft.com

As you can see, what is displayed is dependent upon on the data entered.

The output of virtually every style needs to change depending on whether you have a "Corporate Author" or a "Normal Author." You will see how to use one of the most common rules for implementing such logic into your style, allowing you to display a corporate author if the corporate author is specified, and a normal author if the corporate author is not specified.


### Solution overview

To display a corporate author only if appropriate, use the following procedure.


### To display a corporate author


1. Add a variable to count the number of corporate authors in the citation section of the code.
    
2. Display the corporate author in the citation if the corporate author is filled in. Display the normal author in the citation if the corporate author is not filled in.
    
3. Add a variable to count the number of corporate authors in the bibliography section of the code.
    
4. Display the corporate author in the bibliography if the corporate author is filled in. Display the normal author in the bibliography if the corporate author is not filled in.
    

### Getting started

Let's start by changing the citation. Here is the code for citations from last time.


```XML
<!--Defines the output of the Citation-->
<xsl:template match="b:Citation/b:Source[b:SourceType = 'Book']"> 
   <html xmlns="http://www.w3.org/TR/REC-html40"> 
      <body> 
         <!--Defines the output format as (Author, Year)-->
         <xsl:text>(</xsl:text> 
         <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/>
         <xsl:text>, </xsl:text> 
         <xsl:value-of select="b:Year"/> 
         <xsl:text>)</xsl:text> 
      </body>
   </html> 
</xsl:template>
```


### Step 1: Define a new variable in the citation section to count the number of corporate authors

Declare a new variable to help determine whether a corporate author is available. This variable is a count of the number of times the corporate author field exists in the source.


```
<!--Defines the output of the Citation-->
<html xmlns="http://www.w3.org/TR/REC-html40">
   <!--Count the number of Corporate Authors (can only be 0 or 1)-->
      <xsl:variable name="cCorporateAuthors"> 
         <xsl:value-of select="count(b:Author/b:Author/b:Corporate)" /> 
      </xsl:variable>
```


### Step 2: Verify that the corporate author has been filled in

Verify that the corporate author has been filled in. You can do this by determining if the count of corporate authors is non-zero. If a corporate author exists, display it. If it does not exist, display the normal author.


```XML

<xsl:text>(</xsl:text> 
<xsl:choose>
<!--When the corporate author exists, display the corporate author-->
<xsl:when test ="$cCorporateAuthors!=0"> 
<xsl:value-of select="b:Author/b:Author/b:Corporate"/> 
</xsl:when>
<!-- When the corporate author does not exist, display the normal author--> 
<xsl:otherwise> 
<xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/> 
</xsl:otherwise> 
</xsl:choose> 
<xsl:text>, </xsl:text>
```

Now that you've made the change for citations, make the change for the bibliography. Here's the bibliography section from earlier in this article.




```XML
<!--Defines the output format for a simple Book (in the Bibliography) with important fields defined-->
<xsl: template match="b:Source[b:SourceType = 'Book']">
<!--Label the paragraph as an Office Bibliography paragraph--> 
<p> 
<xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/> 
<xsl:text>, </xsl:text> 
<xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:First"/> 
<xsl:text>. (</xsl:text> 
<xsl:value-of select="b:Year"/> 
<xsl:text>). </xsl:text> 
<i>

```


### Step 3: Define a new variable in the bibliography section

Once again, let's start by adding a counting variable.


```XML
<!--Defines the output format for a simple Book (in the Bibliography) with important fields defined-->
<xsl: template match="b:Source[b:SourceType = 'Book']"> 
<!--Count the number of Corporate Authors (can only be 0 or 1)-->
<xsl:variable name="cCorporateAuthors"> 
<xsl:value-of select="count(b:Author/b:Author/b:Corporate)" /> 
</xsl:variable>
```


### Step 4: Verify that the corporate author has been filled in

Verify that a corporate author exists.


```XML
…..
<xsl:variable name="cCorporateAuthors"> 
<xsl:value-of select="count(b:Author/b:Author/b:Corporate)" /> 
</xsl:variable> 
<p> 
<xsl:choose>
<!--When the corporate author exists display the corporate author-->
<xsl:when test ="$cCorporateAuthors!=0"> 
<xsl:value-of select="b:Author/b:Author/b:Corporate"/> 
<xsl:text>. (</xsl:text> 
</xsl:when> 
<xsl:otherwise> 
<!--When the corporate author does not exist, display the normal author-->
<xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/> 
<xsl:text>, </xsl:text> 
<xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:First"/> 
<xsl:text>. (</xsl:text>
</xsl:otherwise> 
</xsl:choose>
```

Here's the complete final code.




```XML
<?xml version="1.0" ?> 
<!--List of the external resources that we are referencing-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">
   <!--When the bibliography or citation is in your document, it's just HTML-->
   <xsl:output method="html" encoding="us-ascii"/> 
   <!--Match the root element, and dispatch to its children-->
   <xsl:template match="/"> 
      <xsl:apply-templates select="*" /> 
   </xsl:template>
   <!--Set an optional version number for this style-->
   <xsl:template match="b:version"> 
      <xsl:text>2006.5.07</xsl:text> 
   </xsl:template> 
   <!--Defines the name of the style in the References dropdown-->
   <xsl:template match="b:StyleName">     
      <xsl:text>Simple Book Style</xsl:text> 
   </xsl:template> 
   <!--Specifies which fields should appear in the Create Source dialog when in a collapsed state (The Show All Bibliography Fieldscheckbox is cleared)-->
   <xsl:template match="b:GetImportantFields[b:SourceType = 'Book']"> 
      <b:ImportantFields> 
         <b:ImportantField><xsl:text>b:Author/b:Author/b:NameList</xsl:text> </b:ImportantField> 
         <b:ImportantField> <xsl:text>b:Title</xsl:text> </b:ImportantField> 
         <b:ImportantField> <xsl:text>b:Year</xsl:text> </b:ImportantField> 
         <b:ImportantField> <xsl:text>b:City</xsl:text> </b:ImportantField> 
         <b:ImportantField> <xsl:text>b:Publisher</xsl:text> </b:ImportantField> 
      </b:ImportantFields> 
   </xsl:template>
   <!--Defines the output format for a simple Book (in the Bibliography) with important fields defined-->
   <xsl:template match="b:Source[b:SourceType = 'Book']">
   <!--Count the number of Corporate Authors (can only be 0 or 1-->
   <xsl:variable name="cCorporateAuthors">
      <xsl:value-of select="count(b:Author/b:Author/b:Corporate)" />
   </xsl:variable>
   <!--Label the paragraph as an Office Bibliography paragraph-->
   <p>
      <xsl:choose>
         <xsl:when test ="$cCorporateAuthors!=0">
         <!--When the corporate author exists display the corporate author-->
            <xsl:value-of select="b:Author/b:Author/b:Corporate"/>
            <xsl:text>. (</xsl:text>
         </xsl:when>
         <xsl:otherwise>
            <!--When the corporate author does not exist, display the normal author-->
            <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/>
            <xsl:text>, </xsl:text>
            <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:First"/>
            <xsl:text>. (</xsl:text>
         </xsl:otherwise>
      </xsl:choose>
      <xsl:value-of select="b:Year"/>
      <xsl:text>). </xsl:text>
      <i>
         <xsl:value-of select="b:Title"/>
         <xsl:text>. </xsl:text>
      </i> 
         <xsl:value-of select="b:City"/>
         <xsl:text>: </xsl:text>
         <xsl:value-of select="b:Publisher"/>
         <xsl:text>.</xsl:text>
      </p>
   </xsl:template>
   <!--Defines the output of the entire Bibliography-->
   <xsl:template match="b:Bibliography"> 
      <html xmlns="http://www.w3.org/TR/REC-html40"> 
         <body>
            <xsl:apply-templates select ="*">
            </xsl:apply-templates>
         </body>
      </html>
   </xsl:template>
   <!--Defines the output of the Citation-->
   <xsl:template match="b:Citation/b:Source[b:SourceType = 'Book']">
      <html xmlns="http://www.w3.org/TR/REC-html40"> 
         <xsl:variable name="cCorporateAuthors"> 
            <xsl:value-of select="count(b:Author/b:Author/b:Corporate)" /> 
         </xsl:variable> 
         <body> 
         <!--Defines the output format as (Author, Year--> 
            <xsl:text>(</xsl:text>
            <xsl:choose> 
            <!--When the corporate author exists display the corporate author-->
               <xsl:when test ="$cCorporateAuthors!=0">
                  <xsl:value-of select="b:Author/b:Author/b:Corporate"/>
               </xsl:when>
               <!--When the corporate author does not exist, display the normal author-->
               <xsl:otherwise> 
                  <xsl:value-of select="b:Author/b:Author/b:NameList/b:Person/b:Last"/> 
               </xsl:otherwise>
               </xsl:choose>
               <xsl:text>, </xsl:text> 
               <xsl:value-of select="b:Year"/>
               <xsl:text>)</xsl:text> 
            </body> 
         </html>
   </xsl:template>
   <xsl:template match="text()" />
</xsl:stylesheet>
```


## Conclusion
<a name="Biblio_Conclusion"> </a>

This article showed how to create a custom bibliography style in Word, first by creating a simple style, and then by using conditional statements to create a more complex style.


## Additional resources
<a name="Biblio_AdditionalResources"> </a>


-  [What's new for Word 2013 developers](http://msdn.microsoft.com/library/d7de81f7-ac7f-a88c-4765-7e8f8c7df4b4.aspx)
    
-  [Office and Office 365 Developer Blog](http://msdn.microsoft.com/library/http://blogs.msdn.com/b/officedevdocs/.aspx)
    
-  [Word for developers website](http://msdn.microsoft.com/library/http://msdn.microsoft.com/en-us/office/aa905482.aspx.aspx)
    

