---
title: "Свойство Field.Result (издатель)"
keywords: vbapb10.chm6094855
f1_keywords: vbapb10.chm6094855
ms.prod: publisher
api_name: Publisher.Field.Result
ms.assetid: 213e123e-90a7-32b8-1dcf-37da61a8a7e7
ms.date: 06/08/2017
ms.openlocfilehash: 4fb6fcc7c466c8f60a1f4a3e6705e20156ba30d8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldresult-property-publisher"></a>Свойство Field.Result (издатель)

Возвращает **строку** , представляющую результаты указанного поля. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Результат**

 переменная _expression_A, представляющий объект **поля** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере применяется полужирным шрифтом до первого поля в выделение. В этом примере предполагается, что выбран текст или фигуры с текстом в активной публикации.


```vb
Sub GetFieldResults() 
 If Selection.TextRange.Fields.Count > 0 Then 
 MsgBox "The result of the first field is " &; _ 
 Selection.TextRange.Fields(1).Result &; "." 
 End If 
End Sub
```


