---
title: "Свойство TextEffectFormat.Tracking (издатель)"
keywords: vbapb10.chm3735825
f1_keywords: vbapb10.chm3735825
ms.prod: publisher
api_name: Publisher.TextEffectFormat.Tracking
ms.assetid: 9e110e21-be0c-ec49-6bc4-1ff210de141c
ms.date: 06/08/2017
ms.openlocfilehash: c4620e50e9cb6da3d65b5e47570e82ec298e092b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="texteffectformattracking-property-publisher"></a>Свойство TextEffectFormat.Tracking (издатель)

Возвращает или задает **Variant** , указывающее, отслеживания значение, используемое для отображения пространство между символами в диапазоне указанный текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Отслеживание**

 переменная _expression_A, представляет собой объект- **TextEffectFormat** .


## <a name="remarks"></a>Заметки

Допустимые значения — значение **с плавающей запятой** в диапазоне от 0,0 и 5.0 точки. Для свойства значение 0.0 отключает отслеживание. Неопределенное значения возвращаются в виде -2.


## <a name="example"></a>Пример

В этом примере отключается отслеживание во второй материал, задав свойство **отслеживания** нулевое значение.


```vb
Sub DisableTracking() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Tracking = 0.0 
 
End Sub
```


