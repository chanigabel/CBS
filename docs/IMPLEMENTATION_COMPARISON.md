# השוואת מימוש: Web Path vs Direct-Excel Path

## ארכיטקטורה כללית

| | Web Path | Direct-Excel Path |
|---|---|---|
| **כניסה** | HTTP endpoint (`POST /api/workbook/{id}/normalize`) | CLI / `standardizationOrchestrator.process_workbook_json()` |
| **פורמט נתונים** | JSON בזיכרון (`SheetDataset`) | קובץ Excel ישיר (openpyxl) |
| **שינוי קובץ מקור** | לא — קריאה בלבד | כן — מכניס עמודות לגיליון |
| **מצב session** | נשמר בזיכרון (`SessionRecord`) | אין |
| **גיבוי** | אין | יוצר `.backup` לפני עיבוד |
| **קובץ מרכזי** | `webapp/services/standardization_service.py` | `src/excel_standardization/orchestrator.py` |

---

## טבלת השוואה מפורטת לפי פונקציונליות

| תחום | פונקציונליות | Web Path — מה קורה | Web Path — קובץ / פונקציה | Direct-Excel Path — מה קורה | Direct-Excel Path — קובץ / פונקציה | מנוע משותף |
|---|---|---|---|---|---|---|
| **טעינה** | פתיחת קובץ | `load_workbook(data_only=True)` — ערכים בלבד, ללא נוסחאות | `ExcelToJsonExtractor.extract_workbook_to_json` | `load_workbook(data_only=False, keep_vba=is_macro)` — שומר נוסחאות ו-VBA | `standardizationOrchestrator.process_workbook_json` | — |
| **טעינה** | זיהוי כותרת | `detect_table_region` — ניקוד שורות לפי keywords, מחזיר `TableRegion` | `ExcelReader.detect_table_region` | `find_header` — סריקה ישירה לפי מחרוזת (xlPart) | `ExcelReader.find_header` | `ExcelReader` |
| **טעינה** | כותרות ממוזגות | קריאת ערך מהתא הימני-עליון של הטווח | `ExcelReader._is_merged_cell` | ביטול מיזוגים לפני הכנסת עמודות | `standardizationOrchestrator._unmerge_header_area` | — |
| **טעינה** | שורת עזר מספרית | `_is_column_index_row` — מדלג על שורה עם מספרים רצופים | `ExcelReader._is_column_index_row` | `_remove_numeric_helper_row` — מוחק שורה עם מספרים < 100 | `standardizationOrchestrator._remove_numeric_helper_row` | — |
| **טעינה** | קובץ xlsm | נשמר כ-`.xlsm`; ייצוא תמיד כ-`.xlsx` | `UploadService.handle_upload` / `ExportService.export` | `keep_vba=True` — שומר מאקרו; פלט נשמר כ-`.xlsm` | `standardizationOrchestrator.process_workbook_json` | — |
| **שמות** | ניקוי שם | `clean_name` — שפה, סינון תווים, הסרת טוקנים | `standardizationPipeline.apply_name_standardization` → `NameEngine.normalize_name` | `normalize_name` — אותו מנוע | `NameFieldProcessor._process_simple_name_field` → `NameEngine.normalize_name` | `NameEngine`, `TextProcessor` |
| **שמות** | זיהוי pattern | סורק עד 10 שורות ראשונות | `standardizationPipeline.normalize_dataset` → `detect_first/father_name_pattern` | סורק עד 5 שורות ראשונות | `NameFieldProcessor.detect_father_name_pattern` | `NameEngine` |
| **שמות** | הסרת שם משפחה משם פרטי | Stage A (substring) + Stage B (positional) רק אם pattern != NONE | `standardizationPipeline.apply_name_standardization` → `NameEngine.remove_last_name_from_first_name` | Stage A + Stage B — אותה לוגיקה | `NameFieldProcessor._process_simple_name_field` → `NameEngine.remove_last_name_from_first_name` | `NameEngine` |
| **שמות** | הסרת שם משפחה משם אב | Stage A + Stage B עם pattern | `standardizationPipeline.apply_name_standardization` → `NameEngine.remove_last_name_from_father` | Stage A + Stage B עם pattern | `NameFieldProcessor._process_father_name_field` → `NameEngine.remove_last_name_from_father` | `NameEngine` |
| **שמות** | הכנסת עמודה מתוקנת | לא — מוסיף מפתח `first_name_corrected` ל-dict | `standardizationPipeline.apply_name_standardization` | כן — מכניס עמודה " - מתוקן" ישירות לגיליון | `NameFieldProcessor.prepare_output_columns` → `ExcelWriter.prepare_output_column` | — |
| **שמות** | הדגשת שינויים | אין | — | ורוד (`FFFFC7CE`) בתאים שהשתנו | `ExcelWriter.highlight_changed_cells` | — |
| **מגדר** | נורמליזציה | `normalize_gender` → 1 (זכר) / 2 (נקבה); None/ריק → מחזיר original | `standardizationPipeline.apply_gender_standardization` → `GenderEngine.normalize_gender` | `normalize_gender` → 1 / 2 | `GenderFieldProcessor.process_data` → `GenderEngine.normalize_gender` | `GenderEngine` |
| **מגדר** | זיהוי כותרת | keyword matching — "מין" ב-`FIELD_KEYWORDS` | `ExcelReader.detect_columns` | סריקה מדויקת — תא שערכו בדיוק "מין" (xlWhole) | `GenderFieldProcessor.find_headers` | — |
| **מגדר** | כותרות מרובות | לא — מזהה עמודה אחת | `ExcelReader.detect_columns` | כן — מעבד כל כותרת "מין" בגיליון (VBA FindAllHeaders) | `GenderFieldProcessor.find_headers` | — |
| **מגדר** | הכנסת עמודה מתוקנת | לא — מוסיף `gender_corrected` ל-dict | `standardizationPipeline.apply_gender_standardization` | כן — מכניס "מין - מתוקן" לגיליון | `GenderFieldProcessor.prepare_output_columns` | — |
| **מגדר** | הדגשת שינויים | אין | — | ורוד בתאים שהשתנו | `ExcelWriter.highlight_changed_cells` | — |
| **תאריכים** | זיהוי פורמט DDMM/MMDD | תמיד `DDMM` (hardcoded) | `standardizationPipeline._normalize_date_field` | מזהה אוטומטית מהנתונים: אם first > 12 → DDMM | `DateFieldProcessor.detect_date_format_pattern` | — |
| **תאריכים** | פריסת תאריך מפוצל | `parse_from_split_columns` — year/month/day | `standardizationPipeline._normalize_date_field` → `DateEngine.parse_date` | `parse_from_split_columns` — אותה לוגיקה | `DateFieldProcessor._process_date_field` → `DateEngine.parse_date` | `DateEngine` |
| **תאריכים** | פריסת תאריך יחיד | `parse_from_main_value` — ISO, slash, dot, month name | `standardizationPipeline._normalize_date_field` → `DateEngine.parse_date_value` | `parse_from_main_value` — אותה לוגיקה | `DateFieldProcessor._process_date_field` → `DateEngine.parse_date_value` | `DateEngine` |
| **תאריכים** | datetime object בעמודת שנה | מזהה ומטפל כ-main_val | `standardizationPipeline._normalize_date_field` | לא מטופל במפורש | `DateFieldProcessor._normalize_split_value` | — |
| **תאריכים** | כתיבת סטטוס | שומר ב-`birth_date_status` / `entry_date_status` (מחרוזת) | `standardizationPipeline._normalize_date_field` | כותב לתא Excel; צהוב לגיל מעל 100, ורוד לשגיאות | `DateFieldProcessor._process_date_field` → `ExcelWriter.format_cell` | `DateEngine` |
| **תאריכים** | כניסה לפני לידה | **לא נבדק** — `validate_entry_before_birth` קיים אך לא נקרא | `DateEngine.validate_entry_before_birth` (לא מופעל) | נבדק ומוסיף אזהרה לתא סטטוס הכניסה | `standardizationOrchestrator._validate_entry_vs_birth` | `DateEngine` |
| **תאריכים** | הכנסת עמודות מתוקנות | לא — מוסיף `*_corrected` ל-dict | `standardizationPipeline._normalize_date_field` | כן — מכניס 4 עמודות (שנה/חודש/יום/סטטוס) | `DateFieldProcessor.prepare_output_columns` → `ExcelWriter.insert_output_columns` | — |
| **מזהים** | נורמליזציה משולבת | `normalize_identifiers(id, passport)` — שניהם יחד | `standardizationPipeline.apply_identifier_standardization` → `IdentifierEngine.normalize_identifiers` | `normalize_identifiers` — אותה לוגיקה | `IdentifierFieldProcessor.process_data` → `IdentifierEngine.normalize_identifiers` | `IdentifierEngine` |
| **מזהים** | דרישת שתי כותרות | לא — מעבד גם אם רק אחד קיים | `standardizationPipeline.apply_identifier_standardization` | כן — יוצא אם חסרה כותרת ת.ז. **או** דרכון | `IdentifierFieldProcessor.find_headers` | — |
| **מזהים** | כתיבת סטטוס | שומר ב-`identifier_status` (מחרוזת) | `standardizationPipeline.apply_identifier_standardization` | כותב לתא Excel עם עיצוב | `IdentifierFieldProcessor.process_data` → `ExcelWriter.write_column_array` | `IdentifierEngine` |
| **מזהים** | הכנסת עמודות מתוקנות | לא — מוסיף `id_number_corrected`, `passport_corrected` ל-dict | `standardizationPipeline.apply_identifier_standardization` | כן — מכניס 3 עמודות (ת.ז./דרכון/סטטוס) אחרי עמודת דרכון | `IdentifierFieldProcessor.prepare_output_columns` → `ExcelWriter.insert_output_columns` | — |
| **מזהים** | הדגשת שינויים | אין | — | ורוד בתאים שהשתנו | `ExcelWriter.highlight_changed_cells` | — |
| **שורות** | סינון שורות ריקות | מסנן שורות שכל עמודות המקור ריקות | `WorkbookService.get_sheet_data` / `ExportService.visible_rows` | אין סינון בזמן עיבוד; מסנן בייצוא | `ExportEngine.is_valid_data_row` | — |
| **שורות** | הסרת שורת עזר מספרית (תצוגה) | מסיר שורה ראשונה אם כל ערכיה מספריים | `WorkbookService.get_sheet_data` | מוחק שורה מהגיליון לפני עיבוד | `standardizationOrchestrator._remove_numeric_helper_row` | — |
| **שורות** | עמודת מספר סידורי | מזריק `_serial` סינתטי אם אין עמודה מקורית | `derived_columns.apply_derived_columns` | אין | — | — |
| **שורות** | עמודת MosadID | מזריק מ-metadata אם קיים | `derived_columns.apply_derived_columns` | מגיע מ-tracking dict בייצוא | `ExportEngine._export_sheet_from_worksheet` | — |
| **עריכה** | עריכת תא | מעדכן dict בזיכרון + שומר ב-`record.edits` | `EditService.edit_cell` | **לא קיים** | — | — |
| **עריכה** | מחיקת שורות | מסיר מ-`sheet.rows` בזיכרון (all-or-nothing) | `EditService.delete_rows` | **לא קיים** | — | — |
| **עריכה** | replay עריכות לאחר normalize | **לא קיים** — `record.edits` נשמר אך לא מוחל מחדש | `SessionRecord.edits` | — | — | — |
| **ייצוא** | סכמת עמודות | קבועה: 14/15 עמודות לפי סוג גיליון | `ExportService.EXPORT_MAPPING` / `headers_for_sheet` | גמישה: מבוססת על זיהוי עמודות מתוקנות | `ExportEngine.detect_corrected_columns` | — |
| **ייצוא** | שמות גיליונות | ממפה לשמות קנוניים (DayarimYahidim, MeshkeyBayt, AnasheyTzevet) | `ExportService.canonical_sheet_name` | שמות מקוריים מהגיליון | `ExportEngine.SOURCE_SHEET_SPECS` | — |
| **ייצוא** | כיוון RTL | כן — `ws.sheet_view.rightToLeft = True` | `ExportService.export` | לא — שומר כיוון מקורי | — | — |
| **ייצוא** | fallback לשדה מקורי | אין — תא ריק אם `*_corrected` חסר | `ExportService._cell_value` | יש — `pick(corrected, original)` | `ExportEngine._map_row_to_export_fields` | — |
| **ייצוא** | שם קובץ פלט | `{stem}_normalized_{timestamp}.xlsx` | `ExportService.export` | `{stem}_Export.xlsx` (Desktop) | `standardizationOrchestrator.export_vba_parity_workbook_from_json` | — |
| **ייצוא** | פורמט פלט | תמיד `.xlsx` | `ExportService.export` | שומר `.xlsm` לקבצי מאקרו | `standardizationOrchestrator.process_workbook_json` | — |
| **ייצוא** | הדגשות בפלט | אין | `ExportService.export` | יש — ורוד/צהוב בתאים שהשתנו | `ExcelWriter.highlight_changed_cells` / `format_cell` | — |

---

## מנועים משותפים (לוגיקה טהורה — ללא תלות ב-Excel)

| מנוע | קובץ | מה הוא עושה |
|---|---|---|
| `TextProcessor` | `src/excel_standardization/engines/text_processor.py` | ניקוי טקסט: שפה, סינון תווים, הסרת טוקנים, zero-width |
| `NameEngine` | `src/excel_standardization/engines/name_engine.py` | נורמליזציה + הסרת שם משפחה (Stage A/B) + זיהוי pattern |
| `GenderEngine` | `src/excel_standardization/engines/gender_engine.py` | המרה ל-1/2 לפי רשימת patterns |
| `DateEngine` | `src/excel_standardization/engines/date_engine.py` | פריסת תאריכים (split/single/ISO/serial) + business rules |
| `IdentifierEngine` | `src/excel_standardization/engines/identifier_engine.py` | אימות checksum ת.ז. + ניקוי דרכון + סטטוס |

---

## הבדלים קריטיים בהתנהגות

| # | נושא | Web | Direct-Excel |
|---|---|---|---|
| 1 | **כניסה לפני לידה** | לא נבדק | נבדק ומסומן |
| 2 | **זיהוי DDMM/MMDD** | תמיד DDMM | מזהה אוטומטית |
| 3 | **דרישת שתי כותרות מזהים** | לא נדרש | חובה |
| 4 | **הדגשה ורודה/צהובה** | אין | יש |
| 5 | **replay עריכות** | לא מיושם | לא רלוונטי |
| 6 | **fallback לשדה מקורי בייצוא** | אין | יש |
| 7 | **כיוון RTL בייצוא** | יש | אין |
| 8 | **sample לזיהוי pattern שמות** | 10 שורות | 5 שורות |
