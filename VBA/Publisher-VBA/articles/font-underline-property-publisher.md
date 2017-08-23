---
title: "Свойство Font.Underline (издатель)"
keywords: vbapb10.chm5373987
f1_keywords: vbapb10.chm5373987
ms.prod: publisher
api_name: Publisher.Font.Underline
ms.assetid: a01a943e-274d-725e-3f78-aa76c51d5c46
ms.date: 06/08/2017
ms.openlocfilehash: fedc838954d8514c4d12a4e737fe40bb1c2f3432
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontunderline-property-publisher"></a>Свойство Font.Underline (издатель)

Возвращает или задает константой **PbUnderlineType** , указывающую тип подчеркивание выбранного символы указанного шрифта в диапазон текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Подчеркивание**

 переменная _expression_A, представляющий объект **шрифта** .


### <a name="return-value"></a>Возвращаемое значение

PbUnderlineType


## <a name="remarks"></a>Заметки

Значение свойства **Подчеркивание** может иметь одно из **[PbUnderlineType](pbunderlinetype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере форматов символов первая статья пунктирной и толстые линией.


```vb
Sub DashHeavy() 
 
 Application.ActiveDocument.Stories(1).TextRange.Font.Underline = pbUnderlineDashHeavy 
 
End Sub
```


