---
title: "Свойство FillFormat.Pattern (издатель)"
keywords: vbapb10.chm2359558
f1_keywords: vbapb10.chm2359558
ms.prod: publisher
api_name: Publisher.FillFormat.Pattern
ms.assetid: 5b63c81e-b692-92e0-5d72-99c8d4376aff
ms.date: 06/08/2017
ms.openlocfilehash: f3c53132ef16fa4c3a82b7cf14481a6f4ccf5f5e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatpattern-property-publisher"></a>Свойство FillFormat.Pattern (издатель)

Возвращает константу **MsoPatternType** , представляющее шаблон, применяемый к указанным заливки или строки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Шаблон**

 переменная _expression_A, представляет собой объект- **FillFormat** .


## <a name="remarks"></a>Заметки

Значение свойства **шаблон** может иметь одно из ** [MsoPatternType](http://msdn.microsoft.com/library/b95a7e43-329f-b93b-3664-04d8f570c747%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере задает шаблон для указанного фигуры, если фигуры в настоящее время не задан параметр узор заливки. В этом примере предполагается, что по крайней мере один фигуры существует на первой странице active публикации.


```vb
Sub ChangeFillPattern() 
 With ActiveDocument.Pages(1).Shapes(1).Fill 
 If .Pattern < msoPattern10Percent Then 
 .Patterned Pattern:=msoPattern25Percent 
 End If 
 End With 
End Sub
```


