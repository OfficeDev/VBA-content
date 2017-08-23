---
title: "Метод Application.Help (издатель)"
keywords: vbapb10.chm131125
f1_keywords: vbapb10.chm131125
ms.prod: publisher
api_name: Publisher.Application.Help
ms.assetid: 37b51399-5897-4003-a0a9-9829a8adf8ed
ms.date: 06/08/2017
ms.openlocfilehash: 1c918df0eefcf8d8f5b37dfde65ca7180bb4840c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationhelp-method-publisher"></a>Метод Application.Help (издатель)

Сведения о интерактивной справки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Справка** ( **_HelpType_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|HelpType|Обязательное свойство.| **PbHelpType**|Тип справки для отображения.|

## <a name="remarks"></a>Заметки

Параметр HelpType может иметь одно из следующих **PbHelpType** константы, описанные в библиотеке типов, Microsoft Publisher.



|**Константы**|**Описание**|
|:-----|:-----|
| **pbHelp**|Отображает диалоговое окно **справки** .|
| **pbHelpActiveWindow**|Отображение справки, описывающие команды, связанные с активным представлением или области.|
| **pbHelpPSSHelp**| Отображает сведения о поддержке продукта.|
Некоторые из указанных выше констант могут быть недоступны, в зависимости от языка Английский (США, например), который установлен или установлен.


## <a name="example"></a>Пример

В этом примере отображается список разделов для устранения неполадок печати.


```vb
Sub ShowPrintTroubleshooter() 
 Application.Help (HelpType:=pbHelpPrintTroubleshooter) 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

