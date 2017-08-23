---
title: "Свойство Field.Next (издатель)"
keywords: vbapb10.chm6094854
f1_keywords: vbapb10.chm6094854
ms.prod: publisher
api_name: Publisher.Field.Next
ms.assetid: a8f0a246-c55e-715e-3f97-a2f08c383e87
ms.date: 06/08/2017
ms.openlocfilehash: 03e1a5410cd0528e203d86b50630a1b47df7e954
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldnext-property-publisher"></a>Свойство Field.Next (издатель)

Возвращает объект **[поля](field-object-publisher.md)** , представляющий следующему полю в диапазон текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Далее**

 переменная _expression_A, представляющий объект **поля** .


### <a name="return-value"></a>Возвращаемое значение

Поле


## <a name="example"></a>Пример

В этом примере вносятся поле после первого поля в диапазоне указанный текст полужирным шрифтом. Предполагается, что существует по крайней мере два поля в диапазоне указанный текст.


```vb
Sub GoToNextField() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .Fields(1).Next.TextRange.Font.Bold = msoTrue 
End Sub
```


