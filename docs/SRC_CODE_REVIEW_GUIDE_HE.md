# מדריך ביקורת קוד ל-SRC

## תקציר מנהלים
המסמך מתאר את כל קובצי src ואת תפקידם בפועל במערכת הפעילה. הדגש הוא על המסלול הפעיל של Web/API Dataset, ועל ההבדל בין קוד פעיל לבין legacy/compatibility שנשמר מסיבות היסטוריות או תואמות.

## הארכיטקטורה ברמת-על
- המסלול הפעיל הוא Web/API Dataset Pipeline: upload → extraction → standardization → validation → export.
- שכבת src מחולקת ל-I/O, מנועי נרמול, processing helpers, validator, export, מודלי נתונים, ו-utilities.
- המסלול הישן של direct Excel/VBA נשאר קיים לצרכי תאימות או ארכיון, אך אינו מסלול ההפעלה הראשי של ה-Web.

## זרימה מקצה לקצה
- 1. קובץ עולה ל-API/CLI ונשמר ב-session.
- 2. ExcelReader ו-ExcelToJsonExtractor מזהים גיליונות, כותרות, עמודות ושורות.
- 3. StandardizationPipeline מפעיל מנועי שם, מין, מזהים ותאריך וכותב שדות _corrected.
- 4. InstitutionReportValidator מוסיף statuses ברמת שורה ו-workbook.
- 5. ExportService/ExportEngine כותבים את חוברת הייצוא הסופית ומפיקים דוח עיבוד קומפקטי.

## מפת תלות
- webapp/api/process_file.py → webapp/services/upload_service.py (upload/store)
- webapp/api/process_file.py → webapp/services/standardization_service.py (standardize)
- webapp/api/process_file.py → webapp/services/export_service.py (export)
- src/excel_standardization/orchestrator.py → src/excel_standardization/workbook_json_flow.py (delegates active flow)
- src/excel_standardization/workbook_json_flow.py → src/excel_standardization/io_layer/excel_to_json_extractor.py (extract workbook)
- src/excel_standardization/workbook_json_flow.py → src/excel_standardization/processing/standardization_pipeline.py (normalize sheets)
- src/excel_standardization/workbook_json_flow.py → src/excel_standardization/export/export_engine.py (write export workbook)
- src/excel_standardization/processing/standardization_pipeline.py → src/excel_standardization/engines/* (normalize fields)
- src/excel_standardization/processing/standardization_pipeline.py → src/excel_standardization/validation/institution_report_validator.py (post-normalization validation)
- src/excel_standardization/io_layer/excel_to_json_extractor.py → src/excel_standardization/io_layer/excel_reader.py (detect headers/rows)
- src/excel_standardization/io_layer/xls_reader.py → src/excel_standardization/io_layer/excel_to_json_extractor.py (reuse extractor for .xls)
- src/excel_standardization/validation/institution_report_validator.py → src/excel_standardization/services/sheet_name_resolver.py (canonical names)

## פעיל מול legacy
- פעיל: data_types, json_exporter, workbook_json_flow, engines/*, processing/*, validation/*, io_layer/*, sheet_name_resolver, orchestrator (facade), cli (wrapper).
- תואם/legacy: export/export_engine.py, חלק מה-aliasים ב-__init__.py, ומסלולי direct Excel שהושבתו במפורש.
- כללי: __init__.py של חבילות הם markers וייצוא סמלים בלבד.

## כללי עסק מרכזיים
- MosadID הוא שדה ייצוא/דיווח בלבד; חסר מדווח, אבל אין ולידציית מספריות או מינימום תווים במסלול הפעיל.
- SugMosad עדיין מחייב נומריות ומינימום 3 תווים, אך בדיקת מילון סוג מוסד חסרה.
- MisparZehut כולל checksum וכפילויות, אך אין בדיקת מרשם אוכלוסין או מוסדות קשורים.
- תאריכים מפוצלים מתוקנים לרכיבים תקינים בלבד; ערך לא תקין לא מועתק לשדה המתוקן.
- יש הבדל מהותי בין status שמופיע ב-grid לבין הדוח הקומפקטי של processing-report.

## קבצים
### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/__init__.py
- תפקיד במערכת: קובץ חבילה/מודל נתונים
- אחריות עיקרית: מגדיר מבני נתונים, סמלים או חוזי API פנימיים.
- מחלקות/פונקציות מרכזיות: אין
- קלט/פלט: קלט: ערכי Python/JSON; פלט: מבני נתונים או enums.
- לוגיקה עסקית: זהו חוזה נתונים או כלי עזר; אין בו כלל עסק ראשי.
- תלויות/יבוא: .data_types, .orchestrator, .json_exporter
- מי קורא לו: כל שכבות המערכת, במיוחד extraction/processing/validation/export.
- שייך לזרימה: shared data model
- סיווג: package
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/cli.py
- תפקיד במערכת: קובץ orchestration/תאימות
- אחריות עיקרית: מחבר בין הרצה חיצונית לבין המסלול הפעיל או התואם לאחור.
- מחלקות/פונקציות מרכזיות: setup_logging, parse_arguments, validate_file_path, build_output_path, main
- קלט/פלט: קלט: נתיב קובץ/דאטהסט; פלט: קובץ JSON או Excel ותיעוד הרצה.
- לוגיקה עסקית: מגדיר כיצד מסלול חיצוני נכנס לצינור הפעיל או נחסם.
- תלויות/יבוא: argparse, logging, sys, pathlib, datetime, typing
- מי קורא לו: CLI, webapp orchestration, package importers.
- שייך לזרימה: CLI / compatibility
- סיווג: active/compat
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/data_types.py
- תפקיד במערכת: קובץ חבילה/מודל נתונים
- אחריות עיקרית: מגדיר מבני נתונים, סמלים או חוזי API פנימיים.
- מחלקות/פונקציות מרכזיות: SheetDataset, WorkbookDataset, ColumnHeaderInfo ...
- קלט/פלט: קלט: ערכי Python/JSON; פלט: מבני נתונים או enums.
- לוגיקה עסקית: זהו חוזה נתונים או כלי עזר; אין בו כלל עסק ראשי.
- תלויות/יבוא: dataclasses, enum, typing
- מי קורא לו: כל שכבות המערכת, במיוחד extraction/processing/validation/export.
- שייך לזרימה: shared data model
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/engines/__init__.py
- תפקיד במערכת: קובץ חבילה/מודל נתונים
- אחריות עיקרית: מגדיר מבני נתונים, סמלים או חוזי API פנימיים.
- מחלקות/פונקציות מרכזיות: אין
- קלט/פלט: קלט: ערכי Python/JSON; פלט: מבני נתונים או enums.
- לוגיקה עסקית: זהו חוזה נתונים או כלי עזר; אין בו כלל עסק ראשי.
- תלויות/יבוא: .text_processor, .name_engine, .gender_engine, .date_engine, .identifier_engine
- מי קורא לו: כל שכבות המערכת, במיוחד extraction/processing/validation/export.
- שייך לזרימה: shared data model
- סיווג: package
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/engines/date_engine.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: DateEngine
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: datetime, logging, re, typing, ..data_types
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/engines/gender_engine.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: GenderEngine
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: typing
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/engines/identifier_engine.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: IdentifierEngine
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: logging, typing, ..data_types
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/engines/name_engine.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: NameEngine
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: logging, typing, .text_processor, ..data_types
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/engines/text_processor.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: TextProcessor
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: ..data_types, re
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/export/__init__.py
- תפקיד במערכת: שכבת יצוא
- אחריות עיקרית: מייצר פלט ייצוא או מסייע בכתיבת קבצי פלט.
- מחלקות/פונקציות מרכזיות: אין
- קלט/פלט: קלט: WorkbookDataset או workbook; פלט: קובץ Excel/JSON.
- לוגיקה עסקית: שולט במיפוי שדות, סדר כותרות, ושמירת פלט יציב.
- תלויות/יבוא: תלויות פנימיות בלבד
- מי קורא לו: workbook_json_flow, webapp export services, tests.
- שייך לזרימה: export
- סיווג: package
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/export/export_engine.py
- תפקיד במערכת: שכבת יצוא
- אחריות עיקרית: מייצר פלט ייצוא או מסייע בכתיבת קבצי פלט.
- מחלקות/פונקציות מרכזיות: ExportSheetSpec, ExportEngine
- קלט/פלט: קלט: WorkbookDataset או workbook; פלט: קובץ Excel/JSON.
- לוגיקה עסקית: שולט במיפוי שדות, סדר כותרות, ושמירת פלט יציב.
- תלויות/יבוא: __future__, dataclasses, pathlib, typing, openpyxl, openpyxl.worksheet.worksheet
- מי קורא לו: workbook_json_flow, webapp export services, tests.
- שייך לזרימה: export
- סיווג: legacy/compat
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/io_layer/__init__.py
- תפקיד במערכת: שכבת קריאה/חילוץ ל-Excel
- אחריות עיקרית: קורא, מזהה ומחלץ מידע מתוך גליונות Excel.
- מחלקות/פונקציות מרכזיות: אין
- קלט/פלט: קלט: Worksheet/קובץ Excel; פלט: TableRegion, עמודות או SheetDataset.
- לוגיקה עסקית: מכיל heuristics של גילוי כותרות, תתי-כותרות, merged cells וסינון corrected/status columns.
- תלויות/יבוא: .excel_reader, .excel_to_json_extractor
- מי קורא לו: ExcelToJsonExtractor, xls_reader, workbook_json_flow, tests.
- שייך לזרימה: Excel reading / extraction
- סיווג: package
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/io_layer/column_detection.py
- תפקיד במערכת: שכבת קריאה/חילוץ ל-Excel
- אחריות עיקרית: קורא, מזהה ומחלץ מידע מתוך גליונות Excel.
- מחלקות/פונקציות מרכזיות: detect_date_groups, match_field, detect_date_subcolumns, detect_columns
- קלט/פלט: קלט: Worksheet/קובץ Excel; פלט: TableRegion, עמודות או SheetDataset.
- לוגיקה עסקית: מכיל heuristics של גילוי כותרות, תתי-כותרות, merged cells וסינון corrected/status columns.
- תלויות/יבוא: __future__, re, typing, openpyxl.worksheet.worksheet, ..data_types
- מי קורא לו: ExcelToJsonExtractor, xls_reader, workbook_json_flow, tests.
- שייך לזרימה: Excel reading / extraction
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/io_layer/excel_reader.py
- תפקיד במערכת: שכבת קריאה/חילוץ ל-Excel
- אחריות עיקרית: קורא, מזהה ומחלץ מידע מתוך גליונות Excel.
- מחלקות/פונקציות מרכזיות: ExcelReader
- קלט/פלט: קלט: Worksheet/קובץ Excel; פלט: TableRegion, עמודות או SheetDataset.
- לוגיקה עסקית: מכיל heuristics של גילוי כותרות, תתי-כותרות, merged cells וסינון corrected/status columns.
- תלויות/יבוא: re, typing, openpyxl.worksheet.worksheet, ..data_types, .
- מי קורא לו: ExcelToJsonExtractor, xls_reader, workbook_json_flow, tests.
- שייך לזרימה: Excel reading / extraction
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/io_layer/excel_to_json_extractor.py
- תפקיד במערכת: שכבת קריאה/חילוץ ל-Excel
- אחריות עיקרית: קורא, מזהה ומחלץ מידע מתוך גליונות Excel.
- מחלקות/פונקציות מרכזיות: ExcelToJsonExtractor
- קלט/פלט: קלט: Worksheet/קובץ Excel; פלט: TableRegion, עמודות או SheetDataset.
- לוגיקה עסקית: מכיל heuristics של גילוי כותרות, תתי-כותרות, merged cells וסינון corrected/status columns.
- תלויות/יבוא: logging, typing, openpyxl, openpyxl.worksheet.worksheet, openpyxl.utils.exceptions, ..data_types
- מי קורא לו: ExcelToJsonExtractor, xls_reader, workbook_json_flow, tests.
- שייך לזרימה: Excel reading / extraction
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/io_layer/field_matching.py
- תפקיד במערכת: שכבת קריאה/חילוץ ל-Excel
- אחריות עיקרית: קורא, מזהה ומחלץ מידע מתוך גליונות Excel.
- מחלקות/פונקציות מרכזיות: normalize_text, contains_field_keyword, should_ignore_column, looks_like_data_value, find_label_row
- קלט/פלט: קלט: Worksheet/קובץ Excel; פלט: TableRegion, עמודות או SheetDataset.
- לוגיקה עסקית: מכיל heuristics של גילוי כותרות, תתי-כותרות, merged cells וסינון corrected/status columns.
- תלויות/יבוא: __future__, re, typing, openpyxl.worksheet.worksheet
- מי קורא לו: ExcelToJsonExtractor, xls_reader, workbook_json_flow, tests.
- שייך לזרימה: Excel reading / extraction
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/io_layer/merged_cells.py
- תפקיד במערכת: שכבת קריאה/חילוץ ל-Excel
- אחריות עיקרית: קורא, מזהה ומחלץ מידע מתוך גליונות Excel.
- מחלקות/פונקציות מרכזיות: is_merged_cell, get_merged_cell_range
- קלט/פלט: קלט: Worksheet/קובץ Excel; פלט: TableRegion, עמודות או SheetDataset.
- לוגיקה עסקית: מכיל heuristics של גילוי כותרות, תתי-כותרות, merged cells וסינון corrected/status columns.
- תלויות/יבוא: __future__, typing, openpyxl.worksheet.worksheet
- מי קורא לו: ExcelToJsonExtractor, xls_reader, workbook_json_flow, tests.
- שייך לזרימה: Excel reading / extraction
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/io_layer/sheet_access.py
- תפקיד במערכת: שכבת קריאה/חילוץ ל-Excel
- אחריות עיקרית: קורא, מזהה ומחלץ מידע מתוך גליונות Excel.
- מחלקות/פונקציות מרכזיות: find_header, read_column_array, read_cell_value, get_last_row
- קלט/פלט: קלט: Worksheet/קובץ Excel; פלט: TableRegion, עמודות או SheetDataset.
- לוגיקה עסקית: מכיל heuristics של גילוי כותרות, תתי-כותרות, merged cells וסינון corrected/status columns.
- תלויות/יבוא: __future__, typing, openpyxl.worksheet.worksheet, ..data_types
- מי קורא לו: ExcelToJsonExtractor, xls_reader, workbook_json_flow, tests.
- שייך לזרימה: Excel reading / extraction
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/io_layer/table_detection.py
- תפקיד במערכת: שכבת קריאה/חילוץ ל-Excel
- אחריות עיקרית: קורא, מזהה ומחלץ מידע מתוך גליונות Excel.
- מחלקות/פונקציות מרכזיות: detect_table_region, score_header_row, score_subheader_row, is_column_index_row, find_table_columns
- קלט/פלט: קלט: Worksheet/קובץ Excel; פלט: TableRegion, עמודות או SheetDataset.
- לוגיקה עסקית: מכיל heuristics של גילוי כותרות, תתי-כותרות, merged cells וסינון corrected/status columns.
- תלויות/יבוא: __future__, typing, openpyxl.worksheet.worksheet, ..data_types
- מי קורא לו: ExcelToJsonExtractor, xls_reader, workbook_json_flow, tests.
- שייך לזרימה: Excel reading / extraction
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/io_layer/xls_reader.py
- תפקיד במערכת: שכבת קריאה/חילוץ ל-Excel
- אחריות עיקרית: קורא, מזהה ומחלץ מידע מתוך גליונות Excel.
- מחלקות/פונקציות מרכזיות: _XlsCell, _XlsWorksheet, _col_letter, _xlrd_cell_to_python, get_xls_sheet_names, extract_xls_to_workbook_dataset, extract_xls_sheet_to_dataset
- קלט/פלט: קלט: Worksheet/קובץ Excel; פלט: TableRegion, עמודות או SheetDataset.
- לוגיקה עסקית: מכיל heuristics של גילוי כותרות, תתי-כותרות, merged cells וסינון corrected/status columns.
- תלויות/יבוא: __future__, datetime, logging, pathlib, typing, ..data_types
- מי קורא לו: ExcelToJsonExtractor, xls_reader, workbook_json_flow, tests.
- שייך לזרימה: Excel reading / extraction
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/json_exporter.py
- תפקיד במערכת: שכבת יצוא
- אחריות עיקרית: מייצר פלט ייצוא או מסייע בכתיבת קבצי פלט.
- מחלקות/פונקציות מרכזיות: JsonExporter, generate_output_filenames
- קלט/פלט: קלט: WorkbookDataset או workbook; פלט: קובץ Excel/JSON.
- לוגיקה עסקית: שולט במיפוי שדות, סדר כותרות, ושמירת פלט יציב.
- תלויות/יבוא: json, pathlib, typing, datetime, .data_types
- מי קורא לו: workbook_json_flow, webapp export services, tests.
- שייך לזרימה: export
- סיווג: legacy/compat
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/orchestrator.py
- תפקיד במערכת: קובץ orchestration/תאימות
- אחריות עיקרית: מחבר בין הרצה חיצונית לבין המסלול הפעיל או התואם לאחור.
- מחלקות/פונקציות מרכזיות: StandardizationOrchestrator
- קלט/פלט: קלט: נתיב קובץ/דאטהסט; פלט: קובץ JSON או Excel ותיעוד הרצה.
- לוגיקה עסקית: מגדיר כיצד מסלול חיצוני נכנס לצינור הפעיל או נחסם.
- תלויות/יבוא: logging, typing, .engines.date_engine, .engines.gender_engine, .engines.identifier_engine, .engines.name_engine
- מי קורא לו: CLI, webapp orchestration, package importers.
- שייך לזרימה: CLI / compatibility
- סיווג: active/compat
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/processing/__init__.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: אין
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: .standardization_pipeline
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: package
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/processing/date_standardization.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: detect_date_format_pattern, apply_date_standardization, normalize_date_field, date_corrected_components, apply_birth_year_majority_correction
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: __future__, logging, datetime, typing, ..data_types
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/processing/gender_standardization.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: apply_gender_standardization
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: __future__, logging, typing, ..data_types
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/processing/identifier_standardization.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: apply_identifier_standardization
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: __future__, logging, typing, ..data_types
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/processing/name_standardization.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: apply_name_standardization
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: __future__, logging, typing, ..data_types
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/processing/standardization_pipeline.py
- תפקיד במערכת: שכבת נרמול פעילה
- אחריות עיקרית: מנרמל או מעביר ערכי שדה לפורמט עסקי קנוני.
- מחלקות/פונקציות מרכזיות: StandardizationPipeline
- קלט/פלט: קלט: SheetDataset/JsonRow; פלט: שדות מתוקנים, statuses ומטא-דאטה.
- לוגיקה עסקית: מיישם ניקוי, המרה, fallback ותיקוני שדה, כולל cases של ערך ריק/לא תקין.
- תלויות/יבוא: logging, typing, ..data_types, ..engines.name_engine, ..engines.gender_engine, ..engines.date_engine
- מי קורא לו: StandardizationPipeline, workbook_json_flow, tests.
- שייך לזרימה: standardization / normalization
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/schema_validation.py
- תפקיד במערכת: שכבת ולידציה
- אחריות עיקרית: בודק כללי תחום ומפיק statuses/שגיאות.
- מחלקות/פונקציות מרכזיות: _get_schema_path, load_schema, validate_json_row, validate_sheet_dataset, validate_workbook_dataset ...
- קלט/פלט: קלט: שורות מנורמלות ומטא-דאטה; פלט: findings וסטטוס.
- לוגיקה עסקית: מיישם כללי תחום, חלקם חסרים במפורש כאשר מקור המידע החיצוני לא קיים.
- תלויות/יבוא: json, pathlib, typing, .data_types
- מי קורא לו: StandardizationPipeline, webapp services, tests.
- שייך לזרימה: validation
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/services/__init__.py
- תפקיד במערכת: קובץ חבילה/מודל נתונים
- אחריות עיקרית: מגדיר מבני נתונים, סמלים או חוזי API פנימיים.
- מחלקות/פונקציות מרכזיות: אין
- קלט/פלט: קלט: ערכי Python/JSON; פלט: מבני נתונים או enums.
- לוגיקה עסקית: זהו חוזה נתונים או כלי עזר; אין בו כלל עסק ראשי.
- תלויות/יבוא: תלויות פנימיות בלבד
- מי קורא לו: כל שכבות המערכת, במיוחד extraction/processing/validation/export.
- שייך לזרימה: shared data model
- סיווג: package
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/services/sheet_name_resolver.py
- תפקיד במערכת: קובץ מערכת
- אחריות עיקרית: מממש אחריות מקומית בתוך המערכת.
- מחלקות/פונקציות מרכזיות: _normalize_text, resolve_canonical_sheet_name
- קלט/פלט: קלט/פלט פנימי.
- לוגיקה עסקית: לוגיקה מקומית.
- תלויות/יבוא: unicodedata
- מי קורא לו: צרכנים פנימיים.
- שייך לזרימה: utility
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/validation/__init__.py
- תפקיד במערכת: שכבת ולידציה
- אחריות עיקרית: בודק כללי תחום ומפיק statuses/שגיאות.
- מחלקות/פונקציות מרכזיות: אין
- קלט/פלט: קלט: שורות מנורמלות ומטא-דאטה; פלט: findings וסטטוס.
- לוגיקה עסקית: מיישם כללי תחום, חלקם חסרים במפורש כאשר מקור המידע החיצוני לא קיים.
- תלויות/יבוא: .institution_report_validator
- מי קורא לו: StandardizationPipeline, webapp services, tests.
- שייך לזרימה: validation
- סיווג: package
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/validation/institution_report_validator.py
- תפקיד במערכת: שכבת ולידציה
- אחריות עיקרית: בודק כללי תחום ומפיק statuses/שגיאות.
- מחלקות/פונקציות מרכזיות: ValidationResult, RowValidationResult, InstitutionReportValidator, _to_str, _is_numeric_str, _to_int_safe, _get_field, _get_corrected_or_original
- קלט/פלט: קלט: שורות מנורמלות ומטא-דאטה; פלט: findings וסטטוס.
- לוגיקה עסקית: מיישם כללי תחום, חלקם חסרים במפורש כאשר מקור המידע החיצוני לא קיים.
- תלויות/יבוא: __future__, logging, os, dataclasses, datetime, typing
- מי קורא לו: StandardizationPipeline, webapp services, tests.
- שייך לזרימה: validation
- סיווג: active
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

### C:/Users/ch058/OneDrive/שולחן העבודה/python_automation/src/excel_standardization/workbook_json_flow.py
- תפקיד במערכת: קובץ orchestration/תאימות
- אחריות עיקרית: מחבר בין הרצה חיצונית לבין המסלול הפעיל או התואם לאחור.
- מחלקות/פונקציות מרכזיות: extract_workbook, build_pipeline, normalize_sheets, default_export_path, export_vba_parity_workbook_from_json
- קלט/פלט: קלט: נתיב קובץ/דאטהסט; פלט: קובץ JSON או Excel ותיעוד הרצה.
- לוגיקה עסקית: מגדיר כיצד מסלול חיצוני נכנס לצינור הפעיל או נחסם.
- תלויות/יבוא: __future__, pathlib, typing, .engines.date_engine, .engines.gender_engine, .engines.identifier_engine
- מי קורא לו: CLI, webapp orchestration, package importers.
- שייך לזרימה: CLI / compatibility
- סיווג: active/compat
- הערות ביקורת קוד: קובץ קטן או עזר; חשוב בעיקר כחוליה בשרשרת ולא כמקום ללוגיקת-על.

## סיכום ממצאים חוצי-קבצים
### ממצאי מפתח
- המסלול הפעיל הוא דאטהסט, לא direct Excel.
- יש הפרדה ברורה בין validation ב-grid לבין processing-report הקומפקטי.
- חלק מהדרישות ממומשות חלקית בלבד, בעיקר סביב SugMosad, MisparZehut ו-minimum entry age.

### סיכונים
- יש דוקסטרינג/תגובות ישנות שיכולות להטעות, במיוחד סביב GenderEngine ו-tooltip של MosadID בצד ה-UI.
- וריאציות כתיב וקייס בשדות תאריך דורשות זהירות מתמדת.
- קוד legacy נשאר פעיל חלקית לצורך תאימות; צריך להבחין בו מהמסלול הראשי.

### מה להדגיש ב-ביקורת
- למה ההפרדה בין I/O, מנועים, ולידציה וייצוא טובה.
- אילו כללים הם blocking ואילו warning-only.
- אילו בדיקות עוד חסרות במערכת.

## תסריט הסבר ל-ביקורת
- פתח מהמסלול הפעיל: upload → extraction → standardization → validation → export.
- הדגש אילו בדיקות נשארות רק ב-_validation_status ואילו מגיעות לדוח הקומפקטי.
- הסבר מהו active code ומהו legacy/compatibility.
- סיים במיפוי החוסרים העסקיים הגדולים ובאילו קבצים הם יושבים.