---
title: "Метод Shapes.AddWebControl (издатель)"
keywords: vbapb10.chm2162722
f1_keywords: vbapb10.chm2162722
ms.prod: publisher
api_name: Publisher.Shapes.AddWebControl
ms.assetid: 94b54939-9627-6b38-4375-f1c87fc8c4f7
ms.date: 06/08/2017
ms.openlocfilehash: a967e1e1817c20ef0e97caa0ba86e54c0e46fb20
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddwebcontrol-method-publisher"></a>Метод Shapes.AddWebControl (издатель)

Добавляет новый объект **фигуры** , представляющее управления веб-формы для указанной коллекции **фигур** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddWebControl** ( **_Тип_**, **_слева_**, **_в начало_**, **_Width_**, **_Height_**, **_LaunchPropertiesWindow_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Тип|Обязательный| **PbWebControlType**|Указывает тип элемента управления веб-формы для добавления. Если используется pbWebControlWebComponent, возникает ошибка.|
|Слева|Обязательное свойство.| **Variant**|Положение левого края фигуры, представляющей управления веб-формы.|
|Вверх|Обязательное свойство.| **Variant**|Положение верхнего края фигуры, представляющей управления веб-формы.|
|Width|Обязательное свойство.| **Variant**|Ширина формы, представляющее управления веб-формы. Для кнопок этот параметр игнорируется.|
|Height|Обязательное свойство.| **Variant**|Высота фигуры, представляющей управления веб-формы. Для кнопок этот параметр игнорируется.|
|LaunchPropertiesWindow|Необязательный| **Boolean**|Не поддерживается. Значение по умолчанию — **False**; Если этот аргумент имеет значение **True**, возникает ошибка.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Для параметров слева, Top, ширину и высоту числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

При добавлении активной зоны в веб-элемент управления с помощью **pbWebControlHotSpot** константу, URL-адрес указан свойством **[гиперссылки](textrange-hyperlinks-property-publisher.md)** .

 Обратите внимание на то, что свойство **Shape.Fill** , которое возвращает объект **FillFormat** и свойство **Shape.Line** , которое возвращает объект **LineFormat** , не были доступны из активная область формы. Ошибка выполнения возвращается при попытке доступа к этим свойствам из активная область формы.

Параметр типа может быть одной из констант **PbWebControlType** объявлена в библиотеке типов, Microsoft Publisher и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **pbWebControlCheckBox**|Добавление флажка.|
| **pbWebControlCommandButton**|Добавление кнопки команды.|
| **pbWebControlHotSpot**|Добавляет активной зоны. |
| **pbWebControlHTMLFragment**|Добавляет фрагмент HTML-кода.|
| **pbWebControlListBox**|Добавление поля со списком.|
| **pbWebControlMultiLineTextBox**|Добавляет область несколько строк текста.|
| **pbWebControlOptionButton**|Добавление переключателя.|
| **pbWebControlSingleLineTextBox**|Добавление однострочного текстового поля.|
| **pbWebControlWebComponent**|Этот метод не используется.|

## <a name="example"></a>Пример

Следующий пример добавляет управления веб-формы флажок первой страницы публикации active.


```vb
Dim shpCheckBox As Shape 
 
Set shpCheckBox = ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCheckBox, _ 
 Left:=216, Top:=216, _ 
 Width:=18, Height:=18) 

```

Следующий пример добавляет активные области фигуры на странице 4 активных веб-публикации. Во-первых Четырехконечная звезда автофигуры добавляется на страницу. Рассмотрим процедуру активная область добавляется каждый arm звезда с помощью метода **AddWebControl** с типом **pbWebControlHotSpot**. И, наконец гиперссылки добавляется для каждой активной зоны с помощью свойства **гиперссылки** каждой фигуры активной зоны.




```vb
Dim theDoc As Document 
Dim theStar As Shape 
Dim theWC1 As Shape 
Dim theWC2 As Shape 
Dim theWC3 As Shape 
Dim theWC4 As Shape 
 
Set theDoc = ActiveDocument 
Set theStar = theDoc.Pages(4).Shapes.AddShape _ 
 (Type:=msoShape4pointStar, Left:=200, Top:=25, _ 
 Width:=200, Height:=200) 
 
With theDoc.Pages(4).Shapes 
 
 Set theWC1 = .AddWebControl(Type:=pbWebControlHotSpot, _ 
 Left:=280, Top:=25, Width:=40, Height:=80) 
 With theWC1 
 .Hyperlink.Address = "http://www.contoso.com/page1.htm" 
 End With 
 
 Set theWC2 = .AddWebControl(Type:=pbWebControlHotSpot, _ 
 Left:=320, Top:=105, Width:=80, Height:=40) 
 With theWC2 
 .Hyperlink.Address = "http://www.contoso.com/page2.htm" 
 End With 
 
 Set theWC3 = .AddWebControl(Type:=pbWebControlHotSpot, _ 
 Left:=280, Top:=145, Width:=40, Height:=80) 
 With theWC3 
 .Hyperlink.Address = "http://www.contoso.com/page3.htm" 
 End With 
 
 Set theWC4 = .AddWebControl(Type:=pbWebControlHotSpot, _ 
 Left:=200, Top:=105, Width:=80, Height:=40) 
 With theWC4 
 .Hyperlink.Address = "http://www.contoso.com/page4.htm" 
 End With 
End With
```


