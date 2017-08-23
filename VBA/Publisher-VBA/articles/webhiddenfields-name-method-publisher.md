---
title: "Метод WebHiddenFields.Name (издатель)"
keywords: vbapb10.chm3997703
f1_keywords: vbapb10.chm3997703
ms.prod: publisher
api_name: Publisher.WebHiddenFields.Name
ms.assetid: 9dade2c9-6f6b-8686-90fa-a41c8bb6dfa2
ms.date: 06/08/2017
ms.openlocfilehash: 8abb3e4bf93c5cc88c40894bb5e0ea15ce31fef4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webhiddenfieldsname-method-publisher"></a>Метод WebHiddenFields.Name (издатель)

Возвращает **строку** , представляющую имя скрытого поля Web для кнопки Web.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **WebHiddenFields** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Номер индекса скрытых полей.|

### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере создается кнопки команды Web с скрытого поля, а затем отображает имя поля.


```vb
Sub GetHiddenWebFieldName() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, _ 
 Left:=100, Top:=100, Width:=100, _ 
 Height:=36).WebCommandButton.HiddenFields 
 .Add Name:="User", Value:="Power" 
 MsgBox "The name of the first hidden field is " &; .Name(1) 
 End With 
End Sub
```


