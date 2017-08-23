---
title: "Свойство MailMergeMappedDataField.Index (издатель)"
keywords: vbapb10.chm6553604
f1_keywords: vbapb10.chm6553604
ms.prod: publisher
api_name: Publisher.MailMergeMappedDataField.Index
ms.assetid: c590d1af-f845-7e1d-95bc-c65969ebd0ff
ms.date: 06/08/2017
ms.openlocfilehash: 7faa5d6f35077333dcad7feaa7a34871394f6f25
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergemappeddatafieldindex-property-publisher"></a>Свойство MailMergeMappedDataField.Index (издатель)

Возвращает значение типа **Long** , представляющее положение определенного элемента в указанном семействе сайтов. .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Индекс**

 переменная _expression_A, представляет собой объект- **MailMergeMappedDataField** .


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


