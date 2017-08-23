---
title: "Метод Page.ExportEmailHTML (издатель)"
keywords: vbapb10.chm393273
f1_keywords: vbapb10.chm393273
ms.prod: publisher
api_name: Publisher.Page.ExportEmailHTML
ms.assetid: 6257e9b5-26b5-73ae-7d40-50dd0a764488
ms.date: 06/08/2017
ms.openlocfilehash: 681938bac6cb5b6c800c8523db1f7d4084c30908
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pageexportemailhtml-method-publisher"></a>Метод Page.ExportEmailHTML (издатель)

Экспортирует активную страницу публикации в виде HTML-файла.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ExportEmailHTML** ( **_Имя файла_**)

 переменная _expression_A, представляющий объект **Page** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя файла|Обязательное свойство.| **String**|Имя файла для экспорта HTML-код.|

## <a name="remarks"></a>Заметки

Если указано имя существующей HTML-файл, этот файл будет перезаписан.

Этот метод можно использовать только на активной странице публикации.


## <a name="example"></a>Пример

В следующем примере задается первой страницы в документ как активную страницу и экспортируется в файл этой страницы. (Обратите внимание на то, что действительный путь к файлу для работы этого примера необходимо заменить PathToFile.)


```vb
Sub ExportEmail() 
 Dim strFilePath As String 
 strFilePath = "PathToFile" 
 With ActiveDocument.ActiveView 
 .ActivePage = ActiveDocument.Pages(1) 
 .ActivePage.ExportEmailHTML (strFilePath) 
 End With 
 
End Sub
```


