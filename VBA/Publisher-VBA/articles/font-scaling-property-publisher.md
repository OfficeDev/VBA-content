---
title: "Свойство Font.Scaling (издатель)"
keywords: vbapb10.chm5373977
f1_keywords: vbapb10.chm5373977
ms.prod: publisher
api_name: Publisher.Font.Scaling
ms.assetid: 4ff0c484-12f8-38e3-72fd-dfd34507aec1
ms.date: 06/08/2017
ms.openlocfilehash: a4ef27aceb97b3e4741103d7271361422c8158e9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontscaling-property-publisher"></a>Свойство Font.Scaling (издатель)

Возвращает или задает значение **типа Variant** , используемый для масштабирования ширину знаков в диапазон текста в процентном соотношении от текущего размера. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Масштабирование**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Допустимые значения — 0,1 для 600.0, где число представляет процент от текущего размера шрифта. Неопределенное значения возвращаются в виде -2.


## <a name="example"></a>Пример

В этом примере увеличивает ширины текста в вторая статья в 200%. В данном примере для работы вторая статья с текстом должен существовать в активный документ.


```vb
Sub ScaleUp() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Scaling = 200 
 
End Sub
```


