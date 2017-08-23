---
title: "Свойство WebCommandButton.ButtonType (издатель)"
keywords: vbapb10.chm3932178
f1_keywords: vbapb10.chm3932178
ms.prod: publisher
api_name: Publisher.WebCommandButton.ButtonType
ms.assetid: 9ccec0bc-4f0a-9851-0066-05ee1f144c5c
ms.date: 06/08/2017
ms.openlocfilehash: 5df93c6844f58b6d9e0017e59fd3c617282e1ed1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttonbuttontype-property-publisher"></a>Свойство WebCommandButton.ButtonType (издатель)

Возвращает или задает значение константы **PbCommandButtonType** , которое указывает, будет ли кнопки команды Web снимите или отправить данные формы в. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ButtonType**

 переменная _expression_A, представляет собой объект- **WebCommandButton** .


### <a name="return-value"></a>Возвращаемое значение

PbCommandButtonType


## <a name="remarks"></a>Заметки

Значение свойства **ButtonType** может иметь одно из **[PbCommandButtonType](pbcommandbuttontype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере создается новая кнопка отправки команду Web, назначает текст для отображения на поверхность и адрес электронной почты, к которому следует отправить данные формы.


```vb
Sub NewWebForm() 
 With ActiveDocument.Pages.Add(Count:=1, After:=1) 
 With .Shapes.AddWebControl(Type:=pbWebControlCommandButton, _ 
 Left:=72, Top:=72, Width:=72, Height:=36) 
 With .WebCommandButton 
 .ButtonType = pbCommandButtonSubmit 
 .ButtonText = "Send Form:" 
 .EmailAddress = "someone@example.com" 
 End With 
 End With 
 End With 
End Sub
```


