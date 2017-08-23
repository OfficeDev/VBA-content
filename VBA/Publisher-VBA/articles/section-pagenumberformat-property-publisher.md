---
title: "Свойство Section.PageNumberFormat (издатель)"
keywords: vbapb10.chm7405573
f1_keywords: vbapb10.chm7405573
ms.prod: publisher
api_name: Publisher.Section.PageNumberFormat
ms.assetid: 5b64a352-2fd8-9e19-3425-a7984dd67edd
ms.date: 06/08/2017
ms.openlocfilehash: e4e4f12df1bd86970a96963ea489adcf02888e31
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="sectionpagenumberformat-property-publisher"></a>Свойство Section.PageNumberFormat (издатель)

Задает или возвращает константу **PbPageNumberFormat** , представляется форматирование нумерации страниц. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageNumberFormat**

 переменная _expression_A, представляет собой объект **раздела** .


### <a name="return-value"></a>Возвращаемое значение

PbPageNumberFormat


## <a name="remarks"></a>Заметки

Значение свойства **PageNumberFormat** может иметь одно из **[PbPageNumberFormat](pbpagenumberformat-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Доступны не все **PbPageNumberFormat** констант, в зависимости от языков, включенные или установлены.


## <a name="example"></a>Пример

В этом примере добавляется новый раздел в активный документ, задает формат номера страницы в нижний регистр roman и затем задается номер начальной страницы 1.


```vb
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(2) 
With objSection 
 .PageNumberFormat = pbPageNumberFormatLCRoman 
 .PageNumberStart = 1 
End With 

```


