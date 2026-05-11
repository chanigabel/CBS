"""Verification script for date engine conservative parsing fixes."""
from src.excel_standardization.engines.date_engine import DateEngine
from src.excel_standardization.data_types import DateFormatPattern, DateFieldType, DateInput
from src.excel_standardization.processing.date_standardization import date_corrected_components

engine = DateEngine()

all_pass = True

def check(label, condition):
    global all_pass
    status = "PASS" if condition else "FAIL"
    if not condition:
        all_pass = False
    print(f"  [{status}] {label}")

print("=== MUST REJECT: impossible years in split path ===")
cases = [
    (1234567, 5, 15, "year=1234567"),
    (5280, 5, 15, "year=5280"),
    (2229, 5, 15, "year=2229"),
    (9999, 5, 15, "year=9999"),
    (3000, 5, 15, "year=3000"),
]
for yr, mo, dy, label in cases:
    r = engine.parse_from_split_columns(yr, mo, dy)
    r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
    cy, cm, cd = date_corrected_components(r2)
    check(f"{label}: corrected fields empty, status visible",
          cy == "" and cm == "" and cd == "" and r2.status_text != "")

print()
print("=== MUST REJECT: plain int without serial metadata ===")
int_cases = [
    (1234567, "int 1234567"),
    (120201, "int 120201"),
    (36525, "int 36525 (valid serial, no metadata)"),
    (45657, "int 45657 (valid serial, no metadata)"),
    (999999, "int 999999"),
    (888888, "int 888888"),
]
for val, label in int_cases:
    r = engine.parse_date_value(val, DateFormatPattern.DDMM)
    r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
    cy, cm, cd = date_corrected_components(r2)
    check(f"{label}: corrected fields empty, status visible",
          cy == "" and cm == "" and cd == "" and r2.status_text != "")

print()
print("=== MUST REJECT: 5-digit and 7-digit numeric strings ===")
bad_strings = [
    ("12345", "5-digit string"),
    ("1234567", "7-digit string"),
    ("123456789", "9-digit string"),
]
for val, label in bad_strings:
    r = engine.parse_date_value(val, DateFormatPattern.DDMM)
    r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
    cy, cm, cd = date_corrected_components(r2)
    check(f"{label}: corrected fields empty, status visible",
          cy == "" and cm == "" and cd == "" and r2.status_text != "")

print()
print("=== MUST ACCEPT: valid dates ===")
valid_cases = [
    ("14/03/1985", "DD/MM/YYYY string"),
    ("1985-03-14", "ISO string"),
    ("140385", "6-digit DDMMYY"),
    ("14031985", "8-digit DDMMYYYY"),
    ("01/01/2020", "DD/MM/YYYY 2020"),
]
for val, label in valid_cases:
    r = engine.parse_date_value(val, DateFormatPattern.DDMM)
    r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
    cy, cm, cd = date_corrected_components(r2)
    check(f"{label}: year={r2.year}, corrected fields populated",
          cy != "" and cm != "" and cd != "")

print()
print("=== MUST ACCEPT: Excel serial WITH metadata ===")
serial_cases = [
    (36525, "serial 36525 = 2000-01-01"),
    (38353, "serial 38353 = 2004-12-31 approx"),
]
for val, label in serial_cases:
    r = engine.parse_input(DateInput(
        source_kind="single",
        field_type=DateFieldType.BIRTH_DATE,
        raw_value=val,
        source_is_excel_date_serial=True,
    ))
    cy, cm, cd = date_corrected_components(r)
    check(f"{label}: year={r.year}, corrected fields populated",
          cy != "" and cm != "" and cd != "")

print()
print("=== MUST REJECT: Excel serial with impossible result ===")
bad_serials = [
    (9999999, "serial 9999999 (year > ref+1)"),
    (2958466, "serial 2958466 (beyond Excel max)"),
]
for val, label in bad_serials:
    r = engine.parse_input(DateInput(
        source_kind="single",
        field_type=DateFieldType.BIRTH_DATE,
        raw_value=val,
        source_is_excel_date_serial=True,
    ))
    cy, cm, cd = date_corrected_components(r)
    check(f"{label}: corrected fields empty, status visible",
          cy == "" and cm == "" and cd == "" and r.status_text != "")

print()
if all_pass:
    print("ALL CHECKS PASSED")
else:
    print("SOME CHECKS FAILED")
