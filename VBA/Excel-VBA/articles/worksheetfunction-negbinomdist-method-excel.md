---
title: WorksheetFunction.NegBinomDist Method (Excel)
keywords: vbaxl10.chm137196
f1_keywords:
- vbaxl10.chm137196
ms.prod: excel
api_name:
- Excel.WorksheetFunction.NegBinomDist
ms.assetid: 7749759b-4698-6341-c28b-521087731951
ms.date: 06/08/2017
---


# WorksheetFunction.NegBinomDist Method (Excel)

Returns the negative binomial distribution. NEGBINOMDIST returns the probability that there will be number_f failures before the number_s-th success, when the constant probability of a success is probability_s. This function is similar to the binomial distribution, except that the number of successes is fixed, and the number of trials is variable. Like the binomial, trials are assumed to be independent.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [NegBinom_Dist](worksheetfunction-negbinom_dist-method-excel.md) method.

## Syntax

 _expression_ . **NegBinomDist**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number_f - the number of failures.|
| _Arg2_|Required| **Double**|Number_s - the threshold number of successes.|
| _Arg3_|Required| **Double**|Probability_s - the probability of a success.|

### Return Value

Double


## Remarks

For example, you need to find 10 people with excellent reflexes, and you know the probability that a candidate has these qualifications is 0.3. NEGBINOMDIST calculates the probability that you will interview a certain number of unqualified candidates before finding all 10 qualified candidates. 


- Number_f and number_s are truncated to integers.
    
- If any argument is nonnumeric, NEGBINOMDIST returns the #VALUE! error value.
    
- If probability_s < 0 or if probability > 1, NEGBINOMDIST returns the #NUM! error value.
    
- If number_f < 0 or number_s < 1, NEGBINOMDIST returns the #NUM! error value.
    
- The equation for the negative binomial distribution is:
![Formula](images/awfngbin_ZA06051210.gif)where: x is number_f, r is number_s, and p is probability_s. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

