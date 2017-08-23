---
title: "Свойство Plate.Index (издатель)"
keywords: vbapb10.chm2883589
f1_keywords: vbapb10.chm2883589
ms.prod: publisher
api_name: Publisher.Plate.Index
ms.assetid: 7a16bd86-f0c4-d2df-832e-e9a55fed9068
ms.date: 06/08/2017
ms.openlocfilehash: eb7459fc40cd81683cef3b8bdea5545bc728d871
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="plateindex-property-publisher"></a>Свойство Plate.Index (издатель)

Возвращает значение типа **Long** , представляющее положение определенного элемента в указанном семействе сайтов. .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Индекс**

 переменная _expression_A, представляющий объект **формы** .


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


