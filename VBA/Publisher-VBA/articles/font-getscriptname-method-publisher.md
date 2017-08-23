---
title: "Метод Font.GetScriptName (издатель)"
keywords: vbapb10.chm5374000
f1_keywords: vbapb10.chm5374000
ms.prod: publisher
api_name: Publisher.Font.GetScriptName
ms.assetid: 332860de-33fa-7d6a-ac42-28c39856cff7
ms.date: 06/08/2017
ms.openlocfilehash: 050d5d75852ddadcdb1454ee011d9092f2053c31
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontgetscriptname-method-publisher"></a>Метод Font.GetScriptName (издатель)

Возвращает **строку** , представляющую имя начертания шрифта, используемых в диапазон текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GetScriptName** ( **_Скрипт_**)

 переменная _expression_A, представляющий объект **Font** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Сценарий|Обязательное свойство.| **PbFontScriptType**|Имя скрипта.|

### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Параметр скрипт может иметь одно из **[PbFontScriptType](pbfontscripttype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере проверяет начертания шрифта по умолчанию используется для указанного текста диапазона Tahoma и, если это не так, устанавливает ее в качестве начертания шрифта по умолчанию.


```vb
Sub GetScript() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Font 
 If .GetScriptName(Script:=pbFontScriptDefault) <> "Tahoma" Then 
 .SetScriptName Script:=pbFontScriptDefault, _ 
 FontName:="Tahoma" 
 End If 
 End With 
End Sub
```


