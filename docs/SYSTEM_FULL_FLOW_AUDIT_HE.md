# דוח בקרה מלא של זרימת המערכת והוולידציה

## תקציר מנהלים
- המערכת הפעילה היא `Web/API Dataset Pipeline` אחת על בסיס `SessionRecord` משותף.
- יש שלושה רבדים: גריד עם `_validation_status`, דוח עיבוד קומפקטי, ודוח מפורט עם `include_details=true`.
- `MosadID` ו־`SugMosad` מוזרקים לתצוגה ולייצוא גם כשהם מגיעים ממטא־דאטה ולא מן השורה עצמה.
- המערכת מבצעת ניקוי שמות, מזהים ותאריכים, כולל שדות `*_corrected`, סטטוסים, וכפילויות חוברהיות.
- חסרות בדיקות מילון ל־`SugMosad`, בדיקות מרשם/מוסדות קשורים ל־`MisparZehut`, ובדיקת גיל כניסה מינימלי.
- יש תוספות מעבר למסמך: `YomKnisa` מיוצא גם בלשוניות נוספות, יש מיגרציית ת״ז לדרכון, והרחבת שנה דו־ספרתית.
- הדוח הקומפקטי מצומצם יותר מה־UI; לא כל מה שנראה ב־`_validation_status` נכנס לדוח.

## זרימת המערכת

| שלב | מה קורה | קוד | מקום הופעה | כשל / חלקי |
|---|---|---|---|---|
| 1. העלאה | `POST /api/upload` קורא קובץ, בודק גודל, שומר מקור ועותק עבודה. | `webapp/api/upload.py`; `upload_service.py` | `UploadResponse` מחזיר `session_id` ושמות גיליונות. | 413 על גודל; 500 על קובץ לא תקין. |
| 2. Session | נוצר `SessionRecord` עם `status='uploaded'` ו־`workbook_dataset=None`. | `session_service.py`; `models/session.py` | המצב נשמר בזיכרון לכל התהליך. | אם השמירה נכשלת הדוח מסמן שגיאה. |
| 3. זיהוי גיליונות | מדיניות טעינה מרכזית מנתבת `.xlsx/.xlsm` ל־openpyxl ו־`.xls` ל־`xlrd`. | `workbook_loader.py`; `upload_service.py`; `xls_reader.py` | רשימת גיליונות חוזרת למשתמש. | קובץ ישן/פגום מחזיר הודעת שגיאה. |
| 4. טעינת חוברת | אם אין `workbook_dataset`, השירותים טוענים חוברת רק כשצריך דרך אותו loader מרכזי. | `workbook_loader.py`; `workbook_service.py`; `standardization_service.py`; `export_service.py` | הטעינה שקופה למשתמש. | אם אין נתון, מוחזר 404/500. |
| 5. זיהוי טבלה | `ExcelReader` מזהה אזור טבלה, כותרות, ותת־כותרות, ומדלג על עמודות מתוקנות/סטטוסים. | `excel_reader.py`; `table_detection.py`; `column_detection.py` | מיפוי שדות נבנה לפני חילוץ. | בלי כותרות טובות הגיליון מסומן skipped. |
| 6. חילוץ שורות | `ExcelToJsonExtractor` קורא שורה ושומר ערכים כמות שהם או `None`. | `excel_to_json_extractor.py` | שורות נכנסות ל־`SheetDataset`. | תא בודד נכשל לא מפיל את כל הגיליון. |
| 7. הזרקת מטא־דאטה | `MosadID` ו־`SugMosad` מוזרקים לפי session/גיליון/שורות נבחרות. | `derived_columns.py`; `mosad_id_scanner.py`; `institution.py` | הערכים נראים בגריד, ביצוא ובדוח. | חוסר מדווח אך לא חוסם. |
| 8. שמות | `NameEngine` מנקה ומסיר שם משפחה מתוך שם פרטי/שם האב כשצריך. | `name_standardization.py`; `name_engine.py`; `text_processor.py` | נוצרים `*_corrected` בגריד ובייצוא. | כשל שורה נשמר עם הערך המקורי. |
| 9. מגדר | טקסט מוכר ממופה ל־`1`/`2`; ערך ריק נשאר ריק. | `gender_standardization.py`; `gender_engine.py`; `_validate_min` | `gender_corrected` מוצג ומתועד. | ערך לא מוכר מסומן כבעיה. |
| 10. מזהים | ת״ז ודרכון מנוקות; ת״ז לא תקינה יכולה לעבור לדרכון; `חסר מזהים` מסומן. | `identifier_standardization.py`; `identifier_engine.py`; `_validate_mispar_zehut` | `identifier_status` ו־`_validation_status` נראים למשתמש. | כפילות/חוסר מדווחות; מרשם ומוסדות קשורים חסרים. |
| 11. תאריכים | תאריך יחיד או מפוצל; הרחבת שנה דו־ספרתית; בדיקת תאריך קלנדרי אמיתי. | `date_standardization.py`; `date_engine.py` | `*_corrected` ו־`*_date_status` נכתבים. | שנה מוקדמת/עתידית/תאריך לא קיים מסומנים. |
| 12. ולידציה | `InstitutionReportValidator` כותב `_validation_status` על שורה ומבצע גם בדיקות חוברתיות. | `institution_report_validator.py`; `standardization_service.py` | הסטטוס נגיש ב־UI. | שגיאות אינן חוסמות יצוא אלא אם יש כשל מערכת. |
| 13. כפילויות | מאגר מזהים חוברתי מזהה כפילויות בתוך גיליון ובין גיליונות. | `validate_workbook`; `validate_sheet` | כפילות נראית ב־UI ובסטטוס. | הדוח הקומפקטי מצמצם רק מסרים אמיתיים. |
| 14. שדות ייצוא חסרים | נבדק לפי סכמת הייצוא אחרי הזרקת מטא־דאטה. | `report_collectors.py`; `export_schema.py`; `export_rows.py` | נכנס לדוח קומפקטי ול־per-sheet warnings. | חוסר לא עוצר יצוא. |
| 15. סטטוס | `success` / `partial_success` / `failed` לפי שגיאות, אזהרות וסיכומים. | `processing_report_service.py`; `report_status_builder.py` | מופיע ב־API, בכותרות `process-file`, ובדוח. | שגיאה חריפה מחזירה `failed`. |
| 16. דוח | דוח קומפקטי כברירת מחדל; `include_details=true` מוסיף פירוט שורות. | `processing_report.py`; `processing_report_service.py` | נשלח ב־`GET /processing-report`. | ללא דוח מוחזר 404. |
| 17. UI | לשוניות קובץ/גיליון, גריד, overlay, כפתורי ייצוא ומחיקה, ותגי אזהרה. | `index.html`; `upload.js`; `grid.js`; `report.js`; `app.js` | המשתמש רואה גריד ותצוגה מלאה. | כישלון טעינה מציג שגיאה. |
| 18. עריכה | עריכה מקוונת משנה ערכים ושומרת טיפוס מקורי כשאפשר. | `edit_service.py`; `edit.js` | השינוי נשמר ב־`record.edits` ובגריד. | UID חסר או שדה לא קיים מחזירים 4xx. |
| 19. הקצאת מוסד | `mosad_id`, `mosad_name`, ו־`mosad_types` נשמרים בסשן; `SugMosad` יכול להיות workbook/sheet/selected_rows. | `institution.py`; `models/requests.py`; `models/session.py` | הערכים משמשים בייצוא ובדוח. | `SugMosad` לא תקין נדחה. |
| 20. יצוא | המערכת כותבת חוברת חדשה, מנקה קבצי ייצוא ישנים, ומחזירה `FileResponse`. | `export_service.py`; `export_writer.py` | הקובץ נשמר בתיקיית output. | כשל כתיבה מחזיר 500. |
| 21. שדות מתוקנים | שמות, מגדר, מזהים ותאריכים נכתבים בשדות מתוקנים. | `standardization_pipeline.py`; helper modules | הגריד, הייצוא והדוח משתמשים בערך המתוקן. | כשל נשמר יחד עם הערך המקורי ככל האפשר. |
| 22. חוסרים | חוסר בשדה חובה מדווח, אך בדרך כלל לא חוסם. | `report_collectors.py`; `report_status_builder.py` | נראה בדוח הקומפקטי וב־per-sheet warnings. | חסימה קשיחה קיימת רק בשגיאות מערכת. |
| 23. כפילויות | כפילות ת״ז מסומנת בתוך גיליון ובין גיליונות. | `institution_report_validator.py` | נראה בעיקר ב־UI וב־_validation_status. | אין חסימה קשיחה. |
| 24. גיבוי | טעינה עצלה, חיפוש גיליון לפי צורך, ונתיב `.xls` נפרד. | `workbook_service.py`; `xls_reader.py`; `standardization_service.py` | שקוף למשתמש. | אם הגיבוי נכשל, חוזר 404/500. |
| 25. שגיאות ואזהרות | שגיאות מערכת נרשמות, אזהרות נשמרות בדוח, ורק כשל חמור עוצר את הזרימה. | `process_file.py`; `export_service.py`; `processing_report_service.py` | המשתמש רואה שגיאות/סטטוס. | אזהרה בלבד מייצרת `partial_success`. |

## שדות עסקיים

### MosadID

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: חובה בייצוא בלבד; אין במסמך נומריות או מינימום אורך.
- סיווג: ממומש חלקית ומדויק למסמך.
- ממומש בפועל:
- השדה מיוצא בכל שלוש הלשוניות.
- `WorkbookService` ו־`ExportService` מזריקים `MosadID` מה־session או מסריקת הגיליון.
- חוסר מדווח כשדה ייצוא חסר.
- זרימה:
- מקור: `SessionRecord.mosad_id` או `scan_mosad_id`.
- ניקוי: חיתוך רווחים בלבד.
- ולידציה: בודק רק חוסר; אין נומריות או מינימום תווים.
- תצוגה/ייצוא: נראה בגריד ובקובץ המיוצא.
- חוסר: נכנס לדוח בלבד ואינו חוסם.
- מיקומי קוד:
- `webapp/services/export_schema.py`
- `workbook_service.py::get_sheet_data`
- `institution_report_validator.py::_validate_mosad_id`
- פערים:
- אין עוד בדיקות לוגיות מעבר לחוסר.
- מעבר לדרישה:
- טקסט העזרה ב־`index.html` עדיין רומז בטעות על ספרות ואורך.
- הערות:
- זהו שדה ייצוא/דיווח, לא שדה חסימה.

### SugMosad

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: חובה בייצוא; נומרי; מינימום 3; בדיקת מילון חסרה.
- סיווג: ממומש חלקית; הלוגיקה המחמירה קיימת, המילון חסר.
- ממומש בפועל:
- יש validation ב־API וב־UI.
- אפשר להחיל ברמת קובץ/גיליון/שורות נבחרות.
- הערך מוזרק לתצוגה ולייצוא.
- זרימה:
- מקור: `mosad_types[0]` או `SugMosadConfig`.
- ניקוי: נחתך מרווחים ונבדק כנומרי.
- ולידציה: נדרש 3 תווים ומעלה כאשר נשלח.
- דיווח: חוסרים נכנסים ל־processing report.
- פער: אין בדיקת מילון.
- מיקומי קוד:
- `webapp/models/requests.py::_validate_numeric_min3`
- `webapp/api/institution.py::apply_mosad_type_scoped`
- `institution_report_validator.py::_validate_sug_mosad`
- פערים:
- בדיקת מילון סוג מוסד אינה קיימת.
- מעבר לדרישה:
- קיים מיפוי של שורות נבחרות והחלת מספר קבוצות.
- הערות:
- זהו שדה חובה לייצוא, אך לא חסימה קשיחה.

### MisparDiraBeMosad

- שדות / לשוניות: AnasheyTzevet, MeshkeyBayt
- דרישת המסמך: חובה בייצוא בלשוניות הללו; אם קיים חייב להיות נומרי.
- סיווג: חלקי; לא חוסם כשחסר.
- ממומש בפועל:
- השדה קיים בסכמת הייצוא רק בלשוניות הרלוונטיות.
- אם קיים, הוא נבדק כנומרי.
- חוסר מדווח בלבד.
- זרימה:
- מקור: ערך שורה רגיל.
- ניקוי: חיתוך רווחים.
- ולידציה: נומריות בלבד אם קיים.
- תצוגה/ייצוא: מופיע רק בלשוניות הרלוונטיות.
- חוסר: מדווח ואינו חוסם.
- מיקומי קוד:
- `export_schema.py`
- `export_writer.py::write_export_workbook`
- `institution_report_validator.py::_validate_mispar_dira`
- פערים:
- חוסר לא הופך לשגיאה חוסמת.
- מעבר לדרישה:
- `DayarimYahidim` מדלג על הבדיקה.
- הערות:
- יותר שדה ייצוא מאשר שדה חסימה.

### ShemPrati

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: לפי המסמך לא נדרש; הוולידציה לא הוגדרה.
- סיווג: מחמיר מהמסמך.
- ממומש בפועל:
- ניקוי שמות מלא דרך `TextProcessor`.
- הסרת שם משפחה משם פרטי כשזוהה דפוס.
- חוסר מסומן בולידציה.
- זרימה:
- מקור: `first_name`.
- ניקוי: רווחים/תווים מיותרים.
- נורמליזציה: מסיר חלקי שם משפחה.
- ולידציה: נחשב בעיה אם חסר.
- תצוגה/ייצוא: `first_name_corrected`.
- מיקומי קוד:
- `text_processor.py::clean_name`
- `name_engine.py::normalize_first_names`
- `institution_report_validator.py::_validate_shem_prati`
- פערים:
- המערכת מחמירה מהמסמך.
- מעבר לדרישה:
- ניקוי תווים נסתרים וסימני פיסוק.
- הערות:
- הבדיקה נראית ב־UI, בדוח ובייצוא.

### ShemMishpaha

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: לפי המסמך לא נדרש; הוולידציה לא הוגדרה.
- סיווג: מחמיר מהמסמך.
- ממומש בפועל:
- ניקוי שם משפחה מלא.
- מסומן כבעיה אם חסר.
- נכנס ל־`last_name_corrected`.
- זרימה:
- מקור: `last_name`.
- ניקוי: חיתוך תווים מיותרים.
- נורמליזציה: נשמר כערך נקי.
- ולידציה: חסר מסומן כבעיה.
- תצוגה/ייצוא: `last_name_corrected`.
- מיקומי קוד:
- `text_processor.py::clean_name`
- `name_engine.py::normalize_father_names`
- `institution_report_validator.py::_validate_shem_mishpaha`
- פערים:
- מחמיר מהמסמך.
- מעבר לדרישה:
- ניקוי דומה לשם פרטי.
- הערות:
- פעיל ב־UI, בדוח ובייצוא.

### ShemHaAv

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: לפי המסמך לא נדרש; אין ולידציה ייעודית.
- סיווג: חלקי; יש ניקוי וייצוא אך אין ולידציה עצמאית.
- ממומש בפועל:
- שם האב מנוקה ונשמר.
- אין פונקציית ולידציה ייעודית.
- נכנס לייצוא כ־`father_name_corrected`.
- זרימה:
- מקור: `father_name`.
- ניקוי: תווים מיותרים ורווחים.
- נורמליזציה: ערך מתוקן נשמר.
- ולידציה: אין מסלול נפרד.
- תצוגה/ייצוא: השדה מוצג ומיוצא.
- מיקומי קוד:
- `text_processor.py::clean_name`
- `name_engine.py::normalize_father_names`
- `export_schema.py`
- פערים:
- אין בדיקה ייעודית לשדה.
- מעבר לדרישה:
- ניקוי פעיל גם בלי דרישה מפורשת.
- הערות:
- זהו שדה ניקוי/ייצוא יותר משדה ולידציה.

### MisparZehut

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: חובה; ת״ז תקינה; ספרת ביקורת; ייחודיות; מרשם; מוסדות קשורים.
- סיווג: חלקי; יש ספרת ביקורת וכפילויות, אך חסרים מרשם ומוסדות קשורים.
- ממומש בפועל:
- ניקוי ת״ז, בדיקת ספרת ביקורת, ומיגרציית ערכים בעייתיים לדרכון.
- כפילויות נבדקות בתוך גיליון ובין גיליונות.
- חוסר/לא תקין/כפילות נראים ב־`identifier_status` וב־UI.
- זרימה:
- מקור: `id_number` ו־`passport`.
- ניקוי: מסיר מקפים ושומר דרכון באוצר תווים מורחב.
- נורמליזציה: ת״ז לא תקינה יכולה לעבור לדרכון.
- ולידציה: חוסר, ספרת ביקורת, וכפילויות.
- חוצה־שדות: דרכון יכול לקלוט ערך שנכשל בת״ז.
- פער: אין מרשם או מוסדות קשורים.
- מיקומי קוד:
- `identifier_engine.py::normalize_identifiers`
- `identifier_standardization.py::apply_identifier_standardization`
- `institution_report_validator.py::_validate_mispar_zehut`
- פערים:
- אין בדיקת מרשם או מוסדות קשורים.
- מעבר לדרישה:
- ת״ז לא תקינה יכולה לעבור לדרכון.
- הערות:
- כפילות נראית ב־UI; הדוח הקומפקטי מצומצם יותר.

### Darkon

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: לא נדרש; לא הוגדרה ולידציה.
- סיווג: חלקי; יש ניקוי וייצוא, אין ולידציה ייעודית.
- ממומש בפועל:
- `clean_passport` שומר תווים שימושיים.
- השדה נכנס לייצוא.
- יכול לקלוט ערך שנדד מת״ז לא תקינה.
- זרימה:
- מקור: `passport`.
- ניקוי: שומר אותיות, ספרות ומקפים.
- נורמליזציה: נשמר כ־`passport_corrected`.
- ולידציה: אין ולידציה ייעודית.
- תצוגה/ייצוא: מוצג ומיוצא.
- מיקומי קוד:
- `identifier_engine.py::clean_passport`
- `identifier_engine.py::normalize_identifiers`
- `export_writer.py::write_export_workbook`
- פערים:
- אין ולידציה ייעודית.
- מעבר לדרישה:
- קולט ערכים שנדדו מת״ז לא תקינה.
- הערות:
- זהו שדה ניקוי/ייצוא.

### Min

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: קוד מגדר 1/2 בלבד.
- סיווג: ממומש.
- ממומש בפועל:
- טקסטים מוכרים ממופים ל־`1`/`2`.
- קלט ריק נשאר ריק.
- ערך לא מוכר מסומן כבעיה.
- זרימה:
- מקור: `gender`.
- ניקוי: קלט ריק נשאר ריק.
- נורמליזציה: טקסטים מוכרים ל־`1`/`2`.
- ולידציה: רק `1`/`2` תקין.
- תצוגה/ייצוא: `gender_corrected`.
- מיקומי קוד:
- `gender_engine.py::normalize_gender`
- `gender_standardization.py::apply_gender_standardization`
- `institution_report_validator.py::_validate_min`
- פערים:
- אין חסימה על ערך ריק.
- מעבר לדרישה:
- מיפוי טקסטואלי רחב יותר מהמסמך.
- הערות:
- `GenderEngine` מחזיר ריק עבור ערכים ריקים.

### ShnatLida

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: שנת לידה תקינה; מינימום 1906; לא עתידית.
- סיווג: חלקי; יש תאריך מלא, הרחבת שנה דו־ספרתית, ואזהרת גיל חריג.
- ממומש בפועל:
- תאריך מפוצל או יחיד נתמך.
- שנה דו־ספרתית מורחבת.
- שנה מוקדמת/עתידית מסומנת.
- זרימה:
- מקור: `birth_year` או `birth_date`.
- ניקוי: המרה לרכיבים נומריים.
- נורמליזציה: שדות `*_corrected` נכתבים.
- ולידציה: 1906 מינימום, לא עתידית, ותאריך קלנדרי אמיתי.
- חוצה־שדות: נבדק מול `entry_date`.
- מיקומי קוד:
- `date_standardization.py`
- `date_engine.py::parse_date / validate_business_rules`
- `institution_report_validator.py::_validate_birth_date`
- פערים:
- המערכת מחמירה מעבר למסמך.
- מעבר לדרישה:
- אזהרת גיל חריג ותיקון רובי של שנה דו־ספרתית.
- הערות:
- אחד האזורים המחמירים ביותר.

### HodeshLida

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: חודש לידה תקין; נומרי; 1-12.
- סיווג: חלקי; נבדק גם תאריך מלא.
- ממומש בפועל:
- `birth_month_corrected` נכתב.
- ערך לא נומרי או מחוץ לטווח מסומן.
- אם אפשר, נבדק תאריך מלא.
- זרימה:
- מקור: `birth_month`.
- ניקוי: המרה למספר.
- נורמליזציה: נכתב ערך מתוקן.
- ולידציה: 1-12 ותאריך מלא.
- תצוגה/ייצוא: הערך מופיע.
- מיקומי קוד:
- `date_standardization.py::normalize_date_field`
- `date_engine.py::_validate_date`
- `institution_report_validator.py::_validate_birth_date`
- פערים:
- המערכת מחמירה מחובת המסמך.
- מעבר לדרישה:
- בדיקת תאריך מלא.
- הערות:
- חוסר בחודש מדווח גם אם המסמך לא מחייב.

### YomLida

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: יום לידה תקין; נומרי; 1-31 בהתאם לחודש.
- סיווג: חלקי; יש בדיקת יום ובדיקת תאריך מלא.
- ממומש בפועל:
- `birth_day_corrected` נכתב.
- ערך לא נומרי או מחוץ לטווח מסומן.
- תאריך לא קלנדרי יכול לרוקן את הרכיב.
- זרימה:
- מקור: `birth_day`.
- ניקוי: נומריות.
- נורמליזציה: ערך מתוקן.
- ולידציה: 1-31 ותאריך מלא.
- תצוגה/ייצוא: הערך מופיע.
- מיקומי קוד:
- `date_standardization.py::date_corrected_components`
- `date_engine.py::_validate_date`
- `institution_report_validator.py::_validate_birth_date`
- פערים:
- המערכת מחמירה מחובת המסמך.
- מעבר לדרישה:
- בדיקת תאריך מלא מעבר ליום.
- הערות:
- ערך יכול להתרוקן אם רק חלק מהרכיבים נכשל.

### shnatknisa

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: שנת כניסה תקינה; נומרית; לא מעבר לשנת המפקד; לא לפני הלידה; גיל מינימום חסר.
- סיווג: חלקי; בדיקת גיל מינימלי לפי סוג מוסד חסרה.
- ממומש בפועל:
- בדיקת שנה מול שנת המפקד קיימת.
- בדיקה שלא לפני הלידה קיימת.
- `entry_date_status` יכול לצבור אזהרת כניסה לפני לידה.
- זרימה:
- מקור: `entry_year` או `entry_date`.
- ניקוי: נומריות.
- נורמליזציה: שדות מתוקנים.
- ולידציה: לא מעבר לשנת המפקד ולא לפני הלידה.
- פער: גיל מינימום לפי סוג מוסד לא ממומש.
- מיקומי קוד:
- `date_standardization.py`
- `date_engine.py::validate_business_rules / validate_entry_before_birth`
- `institution_report_validator.py::_validate_entry_date`
- פערים:
- בדיקת גיל כניסה מינימלי חסרה.
- מעבר לדרישה:
- סטטוס יכול לשלב הודעת כניסה לפני לידה.
- הערות:
- זהו פער עסקי מהותי.

### Hodeshknisa

- שדות / לשוניות: AnasheyTzevet, DayarimYahidim, MeshkeyBayt
- דרישת המסמך: חודש כניסה תקין; נומרי; 1-12.
- סיווג: חלקי; יש בדיקת חודש ותאריך מלא.
- ממומש בפועל:
- `entry_month_corrected` נכתב.
- ערך לא נומרי או מחוץ לטווח מסומן.
- נבדק כחלק מתאריך כניסה מלא.
- זרימה:
- מקור: `entry_month`.
- ניקוי: המרה למספר.
- נורמליזציה: ערך מתוקן.
- ולידציה: 1-12 ותאריך כניסה מלא.
- ייצוג: יש וריאציית כתיב בין לשוניות.
- מיקומי קוד:
- `date_standardization.py::normalize_date_field`
- `date_engine.py::_validate_date`
- `export_schema.py`
- פערים:
- המערכת מחמירה מחובת המסמך.
- מעבר לדרישה:
- יש וריאציית כתיב בין `Hodeshknisa` ל־`HodeshKnisa`.
- הערות:
- הבדיקה עצמה קיימת, אך הייצוא אינו אחיד לגמרי.

### YomKnisa

- שדות / לשוניות: DayarimYahidim; בייצוא גם AnasheyTzevet ו־MeshkeyBayt
- דרישת המסמך: לפי המסמך נדרש רק ל־DayarimYahidim; יום 1-31 בהתאם לחודש.
- סיווג: חלקי ומורחב מעבר למסמך.
- ממומש בפועל:
- הוולידטור דורש את השדה רק עבור `DayarimYahidim`.
- סכמת הייצוא כוללת אותו גם בלשוניות נוספות.
- הוא נבדק כחלק מתאריך כניסה מלא.
- זרימה:
- מקור: `entry_day`.
- ניקוי: נומריות.
- נורמליזציה: `entry_day_corrected`.
- ולידציה: 1-31 ותאריך מלא.
- פער: ייצוא רחב יותר מהמסמך.
- מיקומי קוד:
- `date_standardization.py::normalize_date_field`
- `institution_report_validator.py::_validate_entry_date`
- `export_schema.py`
- פערים:
- הייצוא רחב מהמסמך.
- מעבר לדרישה:
- מיוצא גם בלשוניות נוספות.
- הערות:
- זהו הבדל יצוא ברור.

## סיכומי פערים

### לוגיקה חסרה
- בדיקת מילון `SugMosad`.
- בדיקת מרשם אוכלוסין ל־`MisparZehut`.
- בדיקת מוסדות קשורים ל־`MisparZehut`.
- בדיקת גיל כניסה מינימלי לפי סוג מוסד.

### מעבר לדרישה
- הזרקת `MosadID`/`SugMosad` לתצוגה גם ממטא־דאטה.
- הסתרת שורות ריקות ושורת עזר מספרית.
- מיגרציית ת״ז לדרכון.
- הרחבת שנה דו־ספרתית.
- בדיקת תאריך קלנדרי אמיתי.
- אזהרת גיל חריג.
- `YomKnisa` בייצוא גם בלשוניות נוספות.

### פערי דרישה
- `ShemPrati` ו־`ShemMishpaha` מחמירים מהמסמך.
- שדות תאריך רבים מחמירים מהמסמך.
- `MosadID` אינו נומרי/3 תווים עוד, אך יש טקסט עזרה ישן ב־UI.
- `YomKnisa` מיוצא מעבר למה שהמסמך דורש.
- הדוח הקומפקטי מצומצם יותר מה־UI.

### הבדלי UI מול backend
- `_validation_status` בגריד מציג פרטים שהדוח הקומפקטי לא מציג.
- `include_details=true` מוסיף פירוט רק לתאריכים ולמזהים.
- עריכות מקוונות נשמרות בשרת, לא רק בממשק.

### סיכוני ביקורת
- סטייה בין מסמך ללוגיקה בשדות השם והתאריך.
- פער בין UI לדוח עלול לבלבל בבדיקות קבלה.
- חוסרים ב־`SugMosad` וב־`MisparZehut` הם חורים עסקיים אמיתיים.
- ווריאציות שמות שדות בייצוא דורשות תיעוד קפדני.
- טקסט עזרה ישן על `MosadID` עלול להטעות.
