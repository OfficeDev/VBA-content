---
title: "Свойство TextRange.Fields (издатель)"
keywords: vbapb10.chm5308469
f1_keywords: vbapb10.chm5308469
ms.prod: publisher
api_name: Publisher.TextRange.Fields
ms.assetid: 01efbcae-b65b-68d9-20b0-6bbee31fd762
ms.date: 06/08/2017
ms.openlocfilehash: 7001b266fb2b75e9b7100252b21f0bf0b71ff54f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangefields-property-publisher"></a>Свойство TextRange.Fields (издатель)

Возвращает объект **поля** , представляющий все поля в диапазоне указанный текст.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Поля**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Fields


## <a name="example"></a>Пример

В этом примере полужирного первого поля в первой фигуры на первой странице active публикации.


```vb
Sub CountFields() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Fields(1).TextRange.Font.Bold = msoTrue 
End Sub
```


