---
title: "Свойство TextRange.ContainingObject (издатель)"
keywords: vbapb10.chm5308465
f1_keywords: vbapb10.chm5308465
ms.prod: publisher
api_name: Publisher.TextRange.ContainingObject
ms.assetid: f15c81b5-d03f-0d83-323b-6ec6f57b4f26
ms.date: 06/08/2017
ms.openlocfilehash: c7870b0f673d7f2ff1697a3e1b50e8f2dc1e8d2d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangecontainingobject-property-publisher"></a>Свойство TextRange.ContainingObject (издатель)

Возвращает **объект** , представляющий объект, который содержит диапазон текста. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ContainingObject**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Object


## <a name="example"></a>Пример

В этом примере возвращается имя object, содержащий указанный текст диапазона.


```vb
Sub NameOfContainingObject() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ContainingObject 
 MsgBox The name of the object containing the text is " &; .Name 
 End With 
End Sub
```


