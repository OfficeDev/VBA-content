---
title: Filtering Items Using a String Comparison
ms.prod: outlook
ms.assetid: 90606142-04a9-8591-ecef-61e2a8c5851c
ms.date: 06/08/2017
---


# Filtering Items Using a String Comparison

This topic describes the support for filtering on a string property using Microsoft Jet syntax and DAV Searching and Locating (DASL) syntax.


## Delimiting Strings and Using Escape Characters

When matching string properties, you can use either a pair of single quotes ('), or a pair of double quotes ("), to delimit a string that is part of the filter. For example, all of the following lines function correctly when the property is of type  **String**:


```
sFilter = "[CompanyName] = 'Microsoft'"

sFilter = "[CompanyName] = " &; Chr(34) &; "Microsoft" &; Chr(34)

```

In specifying a filter in a Jet or DASL query, if you use a pair of single quotes to delimit a string that is part of the filter, and the string contains another single quote or apostrophe, then add a single quote as an escape character before the single quote or apostrophe. Use a similar approach if you use a pair of double quotes to delimit a string. If the string contains a double quote, then add a double quote as an escape character before the double quote. 

For example, in the DASL filter string that filters for the  **Subject** property being equal to the word `can't`, the entire filter string is delimited by a pair of double quotes, and the embedded string  `can't` is delimited by a pair of single quotes. There are three characters that you need to escape in this filter string: the starting double quote and the ending double quote for the property reference of `http://schemas.microsoft.com/mapi/proptag/0x0037001f`, and the apostrophe in the value condition for the word  `can't`. Applying the appropriate escape characters, you can express the filter string as follows: 




```
filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" = 'can''t'"
```

Alternatively, you can use the  `chr(34)` function to represent the double quote (whose ASCII character value is 34) that is used as an escape character. Using the `chr(34)` substitution for a double-quote escape character, you can express the last example as follows:




```
filter = "@SQL= " &; Chr(34) &; "http://schemas.microsoft.com/mapi/proptag/0x0037001f" _
    &; Chr(34) &; " = " &; "'can''t'"
```

Escaping single and double quote characters is also required for DASL queries with the  **ci_startswith** or **ci_phrasematch** operators. For example, the following query performs a phrase match query for `can't` in the message subject:




```
filter = "@SQL=" &; Chr(34) &; "http://schemas.microsoft.com/mapi/proptag/0x0037001E" _
    &; Chr(34) &; " ci_phrasematch " &; "'can''t'"
```

Another example is a DASL filter string that filters for the  **Subject** property being equal to the words `the right stuff`, where the word  `stuff` is enclosed by double quotes. In this case, you must escape the enclosing double quotes as follows:




```
filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" = 'the right ""stuff""'"
```

A different set of escaping rules apply to a property reference for named properties that contain the space, single quote, or double quote characters. If the property reference contains a space, single quote, or double quote character, you must use Universal Resource Locator (URL) escaping in the property reference as follows:



| **Character in Property Reference**| **Escape Character**|
|Space character|%20|
|Double quote|%22|
|Single quote|%27|


For example, you would use the following filter to search for a custom property named  **Mom's "Gift"** that contains the word `pearls`:




```
filter = "@SQL=" &; Chr(34) &; _
    "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/" _
    &; "Mom%27s%20%22Gift%22" &; Chr(34) &; " like '%pearls%'"
```


## String Comparisons Using Jet Syntax

The string comparison that Jet filters support is limited to an equivalence matching. You can filter items based on the value of a string property being equivalent to a specific string, for example, the  **[LastName](contactitem-lastname-property-outlook.md)** property being equal to "Wilson". Note that the comparison is not case sensitive; in the last example, specifying "Wilson" and "wilson" as the comparison string will return the same results.


## String Comparisons Using DASL Syntax

The string comparison that DASL filters support includes equivalence, prefix, phrase, and substring matching. Note that when you filter on the  **Subject** property, prefixes such as "RE: " and "FW: " are ignored. For example,


```
sFilter = "[Subject] = 'cat'"
```

will match both "cat" and "RE: cat".


## Equivalence Matching

Similar to Jet filters, DASL filters perform string equivalence comparison by using the equal (=) operator. The value of the string property must be equivalent to the comparison string, with the exception of prefixes "RE: " and "FW: " as mentioned above.

As an example, the following DASL query creates a filter for company name equals 'Microsoft': 




```
criteria = "@SQL=" &; Chr(34) _
&; "urn:schemas-microsoft-com:office:office#Company" &; Chr(34) _
&; " = 'Microsoft'"
```

As another example, assume that the folder you are searching contains items with the following subjects: 


- Question
    
- Questionable
    
- Unquestionable
    
- RE: Question
    
- The big question
    
The following = restriction: 




```
criteria = "@SQL=" &; Chr(34) _ 
&; "urn:schemas:httpmail:subject" &; Chr(34) _ 
&; " = 'question'"
```

will return the following results:


- Question
    
- RE: Question
    

## Prefix, Phrase, and Substring Matching

DASL supports the matching of prefixes, phrases, and substrings in a string property using content indexer keywords  **ci_startswith** and **ci_phrasematch**, and the keyword  **like**. If a store is indexed, searching with content indexer keywords is more efficient than with  **like**. If your search scenarios include substring matching (which content indexer keywords do not support), use the  **like** keyword in a DASL query.

A DASL query can contain  **ci_startswith** or **ci_phrasematch**, and  **like**, but all string comparisons will be carried out as substring matching.


### ci_startswith

The syntax of  **ci_startswith** is as follows:


```
<PropertySchemaName> ci_startswith <ComparisonString> 

```

where  _PropertySchemaName_ is a valid name of a property referenced by namespace, and _ComparisonString_ is the string used for comparison.

 **ci_startswith** performs a search to match prefixes. It uses tokens (characters, word, or words) in the comparison string to match against the first few characters of any word in the string value of the indexed property. If the comparison string contains multiple tokens, every token in the comarison string must have a prefix match in the indexed property. For example:


- Restricting for "sea" would match "search"
    
- Restricting for "sea" would not match "research"
    
- Restricting for "sea" would match "Subject: the deep blue sea"
    
- Restricting for "law order" would match "law and order" or "law &; order"
    
- Restricting for "law and order" would match "I like the show Law and Order."
    
- Restricting for "law and order" would not match "above the law"
    
- Restricting for "sea creatures" would match "Nova special on sea creatures"
    
- Restricting for "sea creatures" would match "sealife creatures"
    
- Restricting for "sea creatures" would not match "undersea creatures"
    
Using the same example in Equivalence Matching, assume that the folder you are searching contains items with the following subjects: 


- Question
    
- Questionable
    
- Unquestionable
    
- RE: Question
    
- The big question
    
The following  **ci_startswith** restriction:




```
criteria = "@SQL=" &; Chr(34) _ 
&; "urn:schemas:httpmail:subject" &; Chr(34) _ 
&; " ci_startswith 'question'" 
```

will return the following results:


- Question
    
- Questionable
    
- RE: Question
    
- The big question
    

### ci_phrasematch

The syntax of  **ci_phrasematch** is as follows:


```
<PropertySchemaName> ci_phrasematch <ComparisonString> 

```

where  _PropertySchemaName_ is a valid name of a property referenced by namespace and _ComparisonString_ is the string used for comparison.

 **ci_phrasematch** performs a search to match phrases. It uses tokens (characters, word, or words) in the comparison string to match entire words in the string value of the indexed property. Tokens are enclosed in double quotes or parentheses. Each token in the comparison string must have a phrase match, and not a substring or prefix match. If the comparison string contains multiple tokens, every token in the comarison string must have a phrase match. Any word within a multiple word property like **Subject** or **Body** can match; it doesn't have to be the first word. For example:


- Restricting for "cat" would match "cat", "cat box", "black cat"
    
- Restricting for "cat" would match "re: cat is out" 
    
- Restricting for "cat" would not match "catalog", "kittycat"
    
- Restricting for "kitty cat" would match "put the kitty cat out"
    
- Restricting for "kitty cat" would not match "great kitty catalog"
    
Using the same example in Equivalence Matching, assume that the folder you are searching contains items with the following subjects: 


- Question
    
- Questionable
    
- Unquestionable
    
- RE: Question
    
- The big question
    
The following  **ci_phrasematch** restriction:




```
criteria = "@SQL=" &; Chr(34) _ 
&; "urn:schemas:httpmail:subject" &; Chr(34) _ 
&; " ci_phrasematch 'question'" 
```

will return the following results:


- Question
    
- RE: Question
    
- The big question
    

### like

 **like** performs prefix, substring, or equivalence matching. Tokens (characters, word, or word) are enclosed with the % character in a specific way depending on the type of matching:


- 
```
  like '<token>%'
```


    provides prefix matching. For example, restricting for
    
```
  like 'cat%'
```


    would match "cat" and "catalog".
    
- 
```
  like '%<token>%'
```


    provides substring matching. For example, restricting for
    
```
  like '%cat%'
```


    would match "cat", "catalog", "kittycat", "decathalon".
    
- 
```
  like '<token>'
```


    provides equivalence matching. For example, restricting for
    
```
  like 'cat'
```


    would match "cat" and "RE: Cat".
    
Each token can match any part of a word in the string property. If the comparison string contains multiple tokens, every token in the comparison string must have a substring match. Any word within a multiple word property like  **Subject** or **Body** can match; it does not have to be the first word.

Using the same example in Equivalence Matching, assume that the folder you are searching contains items with the following subjects: 


- Question
    
- Questionable
    
- Unquestionable
    
- RE: Question
    
- The big question
    
The following like restriction :




```
criteria = "@SQL=" &; Chr(34) _ 
&; "urn:schemas:httpmail:subject" &; Chr(34) _ 
&; " like '%question%'" 
```

will return the following results: 


- Question
    
- Questionable
    
- Unquestionable
    
- RE: Question
    
- The big question
    

