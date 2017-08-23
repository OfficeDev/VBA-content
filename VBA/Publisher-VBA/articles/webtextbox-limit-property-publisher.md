---
title: "Свойство WebTextBox.Limit (издатель)"
keywords: vbapb10.chm4194309
f1_keywords: vbapb10.chm4194309
ms.prod: publisher
api_name: Publisher.WebTextBox.Limit
ms.assetid: b6bf334e-a610-492a-b316-e8b52d223176
ms.date: 06/08/2017
ms.openlocfilehash: ec7966e7137f1864cc6137ed3d3ffa50a4a5e67e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webtextboxlimit-property-publisher"></a>Свойство WebTextBox.Limit (издатель)

Возвращает или задает типа **Long** , который представляет максимальное число символов, которое можно ввести в элемент управления текстового поля Web. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Ограничение**

 переменная _expression_A, представляет собой объект- **WebTextBox** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Текстовое поле ограничения может быть любое число от 1 до 255 символов. Номера больше, чем 255 приведет к ошибке.


## <a name="example"></a>Пример

В этом примере создается новый элемент управления полем Web текст в активной публикации, задает текст по умолчанию и ограничение количества знаков для текстового поля и указывает, что необходимый элемент управления.


```vb
Sub AddWebTextBoxControl() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlMultiLineTextBox, Left:=72, _ 
 Top:=72, Width:=300, Height:=100).WebTextBox 
 .DefaultText = "Please enter text here." 
 .Limit = 200 
 .RequiredControl = msoTrue 
 End With 
End Sub
```


