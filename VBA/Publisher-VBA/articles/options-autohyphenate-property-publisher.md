---
title: "Свойство Options.AutoHyphenate (издатель)"
keywords: vbapb10.chm1048580
f1_keywords: vbapb10.chm1048580
ms.prod: publisher
api_name: Publisher.Options.AutoHyphenate
ms.assetid: 821d0540-80ec-9f9d-777e-4d2596baf7d7
ms.date: 06/08/2017
ms.openlocfilehash: 707503dabd528caaf3d2452d0365701352d594a7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsautohyphenate-property-publisher"></a>Свойство Options.AutoHyphenate (издатель)

 **Значение true** (по умолчанию) для Microsoft Publisher автоматически переносов текста в текстовых рамках. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutoHyphenate**

 переменная _expression_A, представляющий объект **параметров** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере показано включение автоматической расстановки переносов для издателя и задает дискового пространства правого поля для использования при переносы слов для одного дюйма (72 точки).


```vb
Sub SetHyphenationZone() 
 With Options 
 .AutoHyphenate = True 
 .HyphenationZone = 72 
 End With 
End Sub
```


