
# Filtering Items Using a Comparison with a Keywords Property

 **Last modified:** July 28, 2015

The following discussion uses the  **Categories** property as an example, but can apply as well to any multi-valued string property.

The  **Categories** property of an item is of type **keywords** which can contain multiple values. When being compared to a comparison string in a filter, the **Categories** property behaves like a text string where values are separated by a comma and a space. This is true for filters using Microsoft Jet syntax or DAV Searching and Locating (DASL) syntax.


## Jet Queries

 In a Jet query, you can only perform phrase-matching on a keywords property. You cannot perform starts-with or substring matching with a Jet query. Consider the following criteria for ** [Table.Restrict](ecdd30f6-e12c-8025-3ded-592d2fad2bb8.md)**: 


```
string filter  = "[Categories] = 'Partner'"
```

This Jet query will return rows for items where the  **Categories** property for the item finds a phrase-match for **Partner**. It will return rows for items that are categorized as  **Partner**, for items that are categorized as  **Partner** and **Important**, and for items that are categorized as  **Tier1 Partner**. It will not return rows for items that are categorized only as  **Partnership**.


## DASL Queries

To overcome the limitations of keywords restrictions using the Jet query syntax, use DASL syntax which allows starts-with or substring restrictions. The following criteria string will find all items that contain  **Partner** as a category, as one of the words in a category, and as the beginning part of a word in the category, such as the category **Partnership**: 


```
criteria = "@SQL=" &amp; Chr(34) _ 
&amp; "urn:schemas-microsoft-com:office:office#Keywords" _ 
&amp; Chr(34) &amp; " ci_startswith 'Partner'"
```

You can also use a DASL query for equivalence matching in a multi-valued string property. Consider an example where items have one or more of the following four categories: 


- Book
    
- My Book
    
- Book Review
    
- Bookish
    
The DASL equivalence query:




```
criteria = "@SQL=" &amp; Chr(34) _ 
&amp; "urn:schemas-microsoft-com:office:office#Keywords" &amp; Chr(34) _ 
&amp; " = 'Book'"
```

returns any item that has  **Book** as a category, including those categorized with multiple categories, where **Book** is one of the categories. The query does not return items that do not have **Book** as a category.

If the multi-valued property is added to the  ** [Table](0affaafd-93fe-227a-acee-e09a86cadc20.md)** using a reference by namespace, the format of the values of the property is a variant array. To access these values, parse the elements in the array. Using the last example, this would also allow you to obtain the items that contain exactly **Partner** as a category.

