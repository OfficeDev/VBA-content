---
title: Quotation Marks in Strings
keywords: vbaac10.chm5187927
f1_keywords:
- vbaac10.chm5187927
ms.prod: access
ms.assetid: a25474eb-4451-1cb7-aa05-223bd83cf18c
ms.date: 06/08/2017
---


# Quotation Marks in Strings

In situations where you must construct strings to be concatenated, you may need to embed a string within another string, or a string variable within a string. Situations in which you might need to nest one string within another include:


- When specifying criteria for domain aggregate functions .
    
- When specifying criteria for the  **Find** methods.
    
- When specifying criteria for the  **Filter** or **ServerFilter** property of a form.
    
- When building SQL strings.
    

In all of these instances, Access must pass a string to the Microsoft Jet database engine. When you specify a  _criteria_ argument for a domain aggregate function, for example, Access must evaluate any variables, concatenate them into a string, and then pass the entire string to the Jet database engine.

If you embed a numeric variable, Access evaluates the variable and simply concatenates the value into the string. If the variable is a text string, however, the resulting criteria string will contain a string within a string. A string within a string must be identified by string delimiters. Otherwise, the Jet database engine won't be able to determine which part of the string is the value you want to use.
The string delimiters aren't actually part of the variable itself, but they must be included in the string in the  _criteria_ argument. There are three different ways to construct the string in the _criteria_ argument. Each method results in a _criteria_ argument that looks like one of the following examples.



```
"[LastName] = 'Smith'"
```

- or -



```
"[LastName] = ""Smith"""
```


## Include Single Quotation Marks

You should include single quotation marks in the  _criteria_ argument in such a way that when the value of the variable is concatenated into the string, it will be enclosed within the single quotation marks. For instance, suppose your _criteria_ argument must contain a string variable called . You could construct the _criteria_ argument as in the following example:


```
"[LastName] = '" &; strName &; "'"
```

When the variable is evaluated and concatenated into the  _criteria_ string, the _criteria_ string becomes:




```
"[LastName] = 'Smith'"
```


 **Note**  This syntax does not permit the use of apostrophes (') within the value of the variable itself. If the value of the string variable includes an apostrophe, Access generates a run-time error. If your variable may represent values containing apostrophes, consider using one of the other syntax forms discussed in the following sections.


## Include Double Quotation Marks

You should include double quotation marks within the  _criteria_ argument in such a way so that when the value of the variable is evaluated, it will be enclosed within the quotation marks. Within a string, you must use two sets of double quotation marks to represent a single set of double quotation marks. You could construct the _criteria_ argument as in the following example:


```
"[LastName] = """ &; strName &; """"
```

When the variable is evaluated and concatenated into the  _criteria_ argument, each set of two double quotation marks is replaced by one single quotation mark. The _criteria_ argument becomes:




```
"[LastName] = 'Smith'"
```

This syntax may appear more complicated than the single quotation mark syntax, but it enables you to embed a string that contains an apostrophe within the  _criteria_ argument. It also enables you to nest one or more strings within the embedded string.


## Include a Variable Representing Quotation Marks

You can create a string variable that represents double quotation marks, and concatenate this variable into the  _criteria_ argument along with the value of the variable. The ANSI representation for double quotation marks is Chr$(34); you could assign this value to a string variable called . You could then construct the _criteria_ argument as in the following example:


```
"[LastName] = " &; strQuote &; strName &; strQuote
```

When the variables are evaluated and concatenated into the  _criteria_ argument, the _criteria_ argument becomes:




```
[LastName] = "Smith"
```


