---
title: "Метод Font.SetScriptName (издатель)"
keywords: vbapb10.chm5374001
f1_keywords: vbapb10.chm5374001
ms.prod: publisher
api_name: Publisher.Font.SetScriptName
ms.assetid: f1f2c01e-098c-1afd-0e64-1d563c1ca626
ms.date: 06/08/2017
ms.openlocfilehash: cc4b77a199870ee5f2d87b1ae62639993f80f13f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontsetscriptname-method-publisher"></a>Метод Font.SetScriptName (издатель)

Задает имя скрипта шрифта для использования в диапазон текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetScriptName** ( **_Сценарий_**, **_FontName_**)

 переменная _expression_A, представляющий объект **Font** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Сценарий|Обязательное свойство.| **PbFontScriptType**|Имя скрипта.|
|FontName|Обязательное свойство.| **String**|Имя шрифта.|

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


