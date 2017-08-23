---
title: "Свойство WebCommandButton.HiddenFields (издатель)"
keywords: vbapb10.chm3932177
f1_keywords: vbapb10.chm3932177
ms.prod: publisher
api_name: Publisher.WebCommandButton.HiddenFields
ms.assetid: 187553fb-a4d3-a1fb-f583-49e1d76992ec
ms.date: 06/08/2017
ms.openlocfilehash: 9df1ecd58eea0ccf2e29fb59e1eee03aec462f1b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttonhiddenfields-property-publisher"></a>Свойство WebCommandButton.HiddenFields (издатель)

Возвращает объект **WebHiddenFields** , представляющий скрытых полей Web, подключенного к кнопки Отправить.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Скрытые поля**

 переменная _expression_A, представляет собой объект- **WebCommandButton** .


### <a name="return-value"></a>Возвращаемое значение

WebHiddenFields


## <a name="example"></a>Пример

В этом примере добавляет новый скрытого поля Web новой кнопки Отправить.


```vb
Sub CreateActionWebButton() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36).WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 End With 
 .Item(1).WebCommandButton.HiddenFields.Add _ 
 Name:="User", Value:="PowerUser" 
 End With 
End Sub
```


