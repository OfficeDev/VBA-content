---
title: "Метод Document.PrintOutEx (издатель)"
keywords: vbapb10.chm196755
f1_keywords: vbapb10.chm196755
ms.prod: publisher
api_name: Publisher.Document.PrintOutEx
ms.assetid: f11b6f8b-08a0-28f6-5930-47d684585bef
ms.date: 06/08/2017
ms.openlocfilehash: ad9b88b9d63257967f847c4abaaba264c9000cb3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentprintoutex-method-publisher"></a>Метод Document.PrintOutEx (издатель)

Печатает полностью или частично указанной публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Распечатки** ( **_Из_**, **_Чтобы_**, **_параметр PrintToFile_**, **_копий_**, **_сортировки_**, **_PrintStyle_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|From|Необязательный| **Длинный**|Начальный номер страницы.|
|Чтобы|Необязательный| **Длинный**|Номер конечной страницы.|
|Параметр PrintToFile|Необязательный| **String**|Путь и имя документа на печать в файл.|
|Копий|Необязательный| **Длинный**|Число копий для печати.|
|Сопоставление|Необязательный| **Boolean**|При печати нескольких копий документа, **значение True** для печати всех страниц документа перед печатью следующей копии.|
|PrintStyle|Необязательный| **PbPrintStyle**|Стиль печати для использования. Возможные значения см.|

## <a name="remarks"></a>Заметки

Параметр PrintStyle может иметь одно из **[PbPrintStyle](pbprintstyle-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Если PrintStyle **pbPrintStyleMultipleCopiesPerSheet** или **pbPrintStyleMultiplePagesPerSheet**, Publisher игнорирует любое значение, передаваемого для параметра Разобрать по копиям.


## <a name="example"></a>Пример

В этом примере реализуется печать active публикации.


```vb
Sub PrintActivePublication() 
 ThisDocument.PrintOutEx 
End Sub
```


