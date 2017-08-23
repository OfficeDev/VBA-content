---
title: "Свойство LineFormat.Pattern (издатель)"
keywords: vbapb10.chm3408137
f1_keywords: vbapb10.chm3408137
ms.prod: publisher
api_name: Publisher.LineFormat.Pattern
ms.assetid: ba14b1d1-9c32-a58e-d842-52fc3dc985e8
ms.date: 06/08/2017
ms.openlocfilehash: e5d559ed617d7a75d8f644bb083939641301aa08
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatpattern-property-publisher"></a>Свойство LineFormat.Pattern (издатель)

Возвращает или задает константой **MsoPatternType** , представляющее шаблон, применяемый к указанным заливки или строки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Шаблон**

 переменная _expression_A, представляет собой объект- **LineFormat** .


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


