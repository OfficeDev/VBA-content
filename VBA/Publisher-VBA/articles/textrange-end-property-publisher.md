---
title: "Свойство TextRange.End (издатель)"
keywords: vbapb10.chm5308434
f1_keywords: vbapb10.chm5308434
ms.prod: publisher
api_name: Publisher.TextRange.End
ms.assetid: 594cc4b8-d7fb-4b81-4be7-2d416ae513e2
ms.date: 06/08/2017
ms.openlocfilehash: 53008b70817138291dcd24f32365388dc5d0e723
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeend-property-publisher"></a>Свойство TextRange.End (издатель)

Задает или возвращает значение типа **Long** , представляющее конечного символов выделения или диапазон текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **End**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В этом примере запускается выделение 50 знаков текущего фигуры текстовое поле и заканчивается на fiftieth сотен один символ, а затем текст полужирным шрифтом.


```vb
Sub test2() 
 With Selection.TextRange 
 .Start = 50 
 .End = 150 
 .Font.Bold = msoTrue 
 End With 
End Sub
```


