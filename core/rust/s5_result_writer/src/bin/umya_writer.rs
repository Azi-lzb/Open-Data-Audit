use serde::Deserialize;
use serde_json::Value;
use std::env;
use std::io::{self, Read};
use std::path::Path;
use std::time::Instant;

#[derive(Debug, Deserialize)]
struct Payload {
    sheets: Vec<Table>,
}

#[derive(Debug, Deserialize)]
struct Table {
    title: String,
    headers: Vec<Value>,
    #[serde(default)]
    rows: Vec<Vec<Value>>,
}

fn fail(message: impl AsRef<str>) -> ! {
    eprintln!("{}", message.as_ref());
    std::process::exit(2);
}

fn put_cell(sheet: &mut umya_spreadsheet::Worksheet, row: u32, col: u32, value: &Value) {
    let cell = sheet.cell_mut((col, row));
    if value.is_null() {
        cell.set_value_string("");
    } else if let Some(text) = value.as_str() {
        cell.set_value_string(text);
    } else if let Some(flag) = value.as_bool() {
        cell.set_value_bool(flag);
    } else if let Some(number) = value.as_f64() {
        cell.set_value_number(number);
    } else if let Some(decimal) = value.as_object().and_then(|map| map.get("$decimal")).and_then(Value::as_str) {
        let number = decimal.parse::<f64>().unwrap_or_else(|_| fail(format!("invalid decimal: {decimal}")));
        cell.set_value_number(number);
    } else {
        fail(format!("unsupported cell type at row {row}, column {col}"));
    }
}

fn write_payload(payload: Payload, output: &Path) -> Result<(), String> {
    let mut workbook = umya_spreadsheet::new_file();
    for (sheet_index, table) in payload.sheets.iter().enumerate() {
        let sheet = if sheet_index == 0 {
            let sheet = workbook.sheet_mut(0).map_err(|error| error.to_string())?;
            sheet.set_name(&table.title);
            sheet
        } else {
            workbook.new_sheet(&table.title).map_err(|error| error.to_string())?
        };
        for (col, value) in table.headers.iter().enumerate() {
            put_cell(sheet, 1, col as u32 + 1, value);
        }
        for (row, values) in table.rows.iter().enumerate() {
            for (col, value) in values.iter().enumerate() {
                put_cell(sheet, row as u32 + 2, col as u32 + 1, value);
            }
        }
    }
    umya_spreadsheet::writer::xlsx::write(&workbook, output).map_err(|error| error.to_string())
}

fn main() {
    let mut args = env::args_os().skip(1);
    if args.next().as_deref() != Some(std::ffi::OsStr::new("write-tables")) {
        fail("usage: umya-writer write-tables <output.xlsx>");
    }
    let output = match args.next() {
        Some(value) => value,
        None => fail("missing output path"),
    };
    let mut input = String::new();
    if let Err(error) = io::stdin().read_to_string(&mut input) {
        fail(format!("cannot read table payload: {error}"));
    }
    let payload: Payload = serde_json::from_str(&input).unwrap_or_else(|error| fail(format!("invalid table payload: {error}")));
    let started = Instant::now();
    if let Err(error) = write_payload(payload, Path::new(&output)) {
        fail(error);
    }
    eprintln!("engine_seconds={:.9}", started.elapsed().as_secs_f64());
}
