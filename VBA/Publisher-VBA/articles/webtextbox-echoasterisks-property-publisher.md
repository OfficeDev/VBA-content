---
title: "Свойство WebTextBox.EchoAsterisks (издатель)"
keywords: vbapb10.chm4194308
f1_keywords: vbapb10.chm4194308
ms.prod: publisher
api_name: Publisher.WebTextBox.EchoAsterisks
ms.assetid: eefab42f-9fe7-e77e-50cd-c4b1b35548f1
ms.date: 06/08/2017
ms.openlocfilehash: eb03e56249f9439713a31bc0a31aeb05466c7f7f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webtextboxechoasterisks-property-publisher"></a>Свойство WebTextBox.EchoAsterisks (издатель)

Указывает, следует ли отображать звездочки вместо текст, введенный в элемент управления текстового поля Web. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EchoAsterisks**

 переменная _expression_A, представляющий объект **WebTextBox** .


### <a name="return-value"></a>Возвращаемое значение

MsoTrue


## <a name="remarks"></a>Заметки

Значение свойства **EchoAsterisks** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Отображает текст, введенный в элемент управления текстового поля Web.|
| **msoTrue**| Отображает звездочки вместо текст, введенный в элемент управления текстового поля Web.|

## <a name="example"></a>Пример

В этом примере создается элемент управления текстового поля Web, устанавливает максимальное ограничение в качестве десяти символов, указывает, что запись является обязательным и маскирует запись с помощью звездочки, когда пользователь вводит в элементе управления.


```vb
Sub AddPasswordTextBox() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlSingleLineTextBox, Left:=100, _ 
 Top:=100, Width:=72, Height:=15) 
 .Name = "Password" 
 With .WebTextBox 
 .Limit = 10 
 .EchoAsterisks = msoTrue 
 .RequiredControl = msoTrue 
 End With 
 End With 
End Sub
```


