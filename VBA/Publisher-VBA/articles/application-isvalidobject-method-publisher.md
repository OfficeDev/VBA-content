---
title: "Метод Application.IsValidObject (издатель)"
keywords: vbapb10.chm131126
f1_keywords: vbapb10.chm131126
ms.prod: publisher
api_name: Publisher.Application.IsValidObject
ms.assetid: 56b2bc3a-3e8e-058c-046a-146f0fbb294a
ms.date: 06/08/2017
ms.openlocfilehash: 2f047a175324908e2ce4fd8a87f22e74c7d56a63
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationisvalidobject-method-publisher"></a>Метод Application.IsValidObject (издатель)

Определяет, является ли указанный объект переменной ссылается на допустимый объект и возвращает **логическое** значение: **True** указанной переменной, которая ссылается на объект является допустимым, если **значение False,** Если был удален объект ссылается переменная.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsValidObject** ( **_Объект_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Объект|Обязательное свойство.| **Object**|Переменная, которая ссылается на объект.|

### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере форматирует строку допустимого объекта.


```vb
Sub ValidShape(shpObject As Shape) 
 
 If Application.IsValidObject object:=shpObject) = True Then 
 With shpObject.Line 
 .DashStyle = msoLineRoundDot 
 .ForeColor.RGB = RGB(Red:=158, Green:=50, Blue:=208) 
 .Weight = 5 
 End With 
 End If 
 
End Sub
```

Используйте следующие процедуры для вызова подпрограмму выше.




```vb
Sub CallValidShape() 
 Call ValidShape(shpObject:=ActiveDocument.Pages(1).Shapes(2)) 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

