---
title: "Свойство WebCommandButton.PostFormData (издатель)"
keywords: vbapb10.chm3932176
f1_keywords: vbapb10.chm3932176
ms.prod: publisher
api_name: Publisher.WebCommandButton.PostFormData
ms.assetid: d04e3172-0d96-856f-af63-341031d92291
ms.date: 06/08/2017
ms.openlocfilehash: 998236da0d1b9d15fa985232cb447a0aa009afbd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttonpostformdata-property-publisher"></a>Свойство WebCommandButton.PostFormData (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, использует ли указанный элемент управления кнопки команды Web метода **Post** или Microsoft Visual Basic, **Получение** при отправке данных формы на веб-сервере. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PostFormData**

 переменная _expression_A, представляет собой объект- **WebCommandButton** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства может быть одной из констант **MsoTriState** объявлена в библиотеке типов, Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Элемент управления использует метод Visual Basic **получить** для отправки данных формы.|
| **msoTrue**|Элемент управления использует метод Visual Basic **Post** для отправки данных формы. Значение по умолчанию.|
Это свойство игнорируется для **сброса** кнопок.


## <a name="example"></a>Пример

В этом примере создается кнопки Отправить форму Web и задает путь и имя скрипта для запуска при нажатии кнопки. В примере также указывается, что веб-форму следует использовать метод Visual Basic **получить** для отправки данных формы.


```vb
Dim shpNew As Shape 
 
Set shpNew = ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCommandButton, Left:=150, _ 
 Top:=150, Width:=75, Height:=36) 
 
With shpNew.WebCommandButton 
 .ButtonText = "Submit" 
 .ButtonType = pbCommandButtonSubmit 
 .ActionURL = "http://www.tailspintoys.com/" _ 
 &; "scripts/ispscript.cgi" 
 .PostFormData = msoFalse 
End With
```


