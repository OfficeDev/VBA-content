---
title: "Свойство MailMergeFilterCriterion.Index (издатель)"
keywords: vbapb10.chm6815745
f1_keywords: vbapb10.chm6815745
ms.prod: publisher
api_name: Publisher.MailMergeFilterCriterion.Index
ms.assetid: e66e5afd-db28-cd00-9692-3b1a6d557198
ms.date: 06/08/2017
ms.openlocfilehash: 6578f9c77f4e93a779fe973eaa7076b8ac0f9cfb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergefiltercriterionindex-property-publisher"></a>Свойство MailMergeFilterCriterion.Index (издатель)

Возвращает значение типа **Long** , представляющее положение определенного элемента в указанном семействе сайтов. .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Индекс**

 переменная _expression_A, представляет собой объект- **MailMergeFilterCriterion** .


## <a name="example"></a>Пример

В следующем примере коллекции **MailMergeDataFields** и отображает **индекса** и **имя** свойства для каждого поля.


```vb
Dim mmfLoop As MailMergeDataField 
 
With ActiveDocument.MailMerge.DataSource 
 If .DataFields.Count > 0 Then 
 For Each mmfLoop In .DataFields 
 Debug.Print "Field " &; mmfLoop.Name _ 
 &; " / Index " &; mmfLoop.Index 
 Next mmfLoop 
 Else 
 Debug.Print "No fields to report." 
 End If 
End With
```

В следующем примере коллекции **формы** и отображает **индекса** и **имя** свойства для каждой формы.




```vb
Dim plaLoop As Plate 
 
If ActiveDocument.Plates.Count > 0 Then 
 For Each plaLoop In ActiveDocument.Plates 
 Debug.Print "Plate " &; plaLoop.Name _ 
 &; " / Index " &; plaLoop.Index 
 Next plaLoop 
Else 
 Debug.Print "No plates to report." 
End If
```


