---
title: "Свойство MailMergeDataField.Index (издатель)"
keywords: vbapb10.chm6422529
f1_keywords: vbapb10.chm6422529
ms.prod: publisher
api_name: Publisher.MailMergeDataField.Index
ms.assetid: f70d0266-0527-6871-632d-b45b617d75d4
ms.date: 06/08/2017
ms.openlocfilehash: edd828953d1034a7cfecc360e0a650e2e33dba81
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafieldindex-property-publisher"></a>Свойство MailMergeDataField.Index (издатель)

Возвращает значение типа **Long** , представляющее положение определенного элемента в указанном семействе сайтов. .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Индекс**

 переменная _expression_A, представляет собой объект- **MailMergeDataField** .


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


