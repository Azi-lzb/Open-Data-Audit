use rust_xlsxwriter::{Color, Format, FormatAlign, Workbook};
use serde::Deserialize;
use serde_json::Value;
use std::collections::HashMap;
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
    widths: Vec<f64>,
    #[serde(default)]
    rows: Vec<Vec<Value>>,
    #[serde(default)]
    formats: HashMap<String, String>,
    #[serde(default)]
    hyperlinks: HashMap<String, String>,
    #[serde(default)]
    numeric_columns: Vec<u32>,
    #[serde(default)]
    text_columns: Vec<u32>,
}

fn fail(message: impl AsRef<str>) -> ! {
    eprintln!("{}", message.as_ref());
    std::process::exit(2);
}

fn decimal_marker(value: &Value) -> Option<&str> {
    value.as_object()?.get("$decimal")?.as_str()
}

fn openpyxl_safe_number(value: f64) -> f64 {
    // Match OpenPyXL's safe_string(float) precision used by S2 Direct OOXML.
    // Writing the raw f64 with rust_xlsxwriter can expose an extra binary-tail
    // digit (for example 1851.8400000000001 instead of 1851.84).
    format!("{value:.15e}").parse::<f64>().unwrap_or(value)
}

fn write_cell(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    value: &Value,
    format: &Format,
) -> Result<(), String> {
    if value.is_null() {
        worksheet.write_blank(row, col, format).map_err(|error| error.to_string())?;
    } else if let Some(text) = value.as_str() {
        worksheet.write_string_with_format(row, col, text, format).map_err(|error| error.to_string())?;
    } else if let Some(flag) = value.as_bool() {
        worksheet.write_boolean_with_format(row, col, flag, format).map_err(|error| error.to_string())?;
    } else if let Some(number) = value.as_f64() {
        worksheet.write_number_with_format(row, col, openpyxl_safe_number(number), format).map_err(|error| error.to_string())?;
    } else if let Some(decimal) = decimal_marker(value) {
        let number = decimal.parse::<f64>().map_err(|_| format!("invalid decimal value: {decimal}"))?;
        let number = openpyxl_safe_number(number);
        worksheet.write_number_with_format(row, col, number, format).map_err(|error| error.to_string())?;
    } else {
        return Err(format!("unsupported cell type at row {}, column {}", row + 1, col + 1));
    }
    Ok(())
}

fn cell_format(table: &Table, row: u32, col: u32, header_format: &Format, cache: &mut HashMap<String, Format>) -> Format {
    if row == 0 {
        return header_format.clone();
    }
    let coordinate = format!("{},{}", row + 1, col + 1);
    let number_format = if table.text_columns.contains(&(col + 1)) {
        "@".to_string()
    } else {
        table.formats.get(&coordinate).cloned().unwrap_or_else(|| "General".to_string())
    };
    let right = table.numeric_columns.contains(&(col + 1));
    let key = format!("{number_format}|{right}");
    cache.entry(key).or_insert_with(|| {
        let mut format = Format::new()
            .set_font_name("微软雅黑")
            .set_font_size(10)
            .set_font_color(Color::RGB(0x222222))
            .set_text_wrap();
        if number_format != "General" {
            format = format.set_num_format(&number_format);
        }
        if right {
            format = format.set_align(FormatAlign::Right);
        }
        format
    }).clone()
}

fn write_payload(payload: Payload, output_path: &Path) -> Result<(), String> {
    let mut workbook = Workbook::new();
    // Preserve the S5 prototype's level-1 default; allow fair per-system benchmarks
    // to request the same compression level as the Python writer.
    let compression_level = env::var("AUDIT_RUST_ZIP_LEVEL")
        .ok()
        .and_then(|value| value.parse::<u8>().ok())
        .filter(|value| *value <= 9)
        .unwrap_or(1);
    workbook.set_zip_compression_level(compression_level);
    let header_format = Format::new()
        .set_font_name("微软雅黑")
        .set_font_size(10)
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color(Color::RGB(0x1F4E78))
        .set_align(FormatAlign::Center)
        .set_text_wrap();

    for table in payload.sheets {
        if table.title.chars().count() > 31 {
            return Err(format!("worksheet name exceeds Excel limit: {}", table.title));
        }
        let column_count = table.headers.len();
        let total_rows = table.rows.len() + 1;
        let last_row = total_rows.saturating_sub(1) as u32;
        let last_col = column_count.saturating_sub(1) as u16;
        let mut worksheet = workbook.add_worksheet();
        worksheet.set_name(&table.title).map_err(|error| error.to_string())?;
        worksheet.set_screen_gridlines(false);
        worksheet.set_freeze_panes(1, 0).map_err(|error| error.to_string())?;
        worksheet.set_paper_size(8).set_landscape().set_print_fit_to_pages(1, 0);
        worksheet.set_repeat_rows(0, 0).map_err(|error| error.to_string())?;
        if column_count > 0 {
            worksheet.autofilter(0, 0, last_row, last_col).map_err(|error| error.to_string())?;
        }
        for (index, width) in table.widths.iter().enumerate() {
            worksheet.set_column_width(index as u16, *width).map_err(|error| error.to_string())?;
        }
        let mut cache = HashMap::new();
        for row in 0..total_rows {
            worksheet.set_row_height(row as u32, if row == 0 { 30.0 } else { 54.0 }).map_err(|error| error.to_string())?;
            let cells = if row == 0 { &table.headers } else { &table.rows[row - 1] };
            for col in 0..column_count {
                let format = cell_format(&table, row as u32, col as u32, &header_format, &mut cache);
                let value = cells.get(col).unwrap_or(&Value::Null);
                let address = format!("{}{}", column_letters(col as u32), row + 1);
                if row > 0 {
                    if let Some(target) = table.hyperlinks.get(&address) {
                        // rust_xlsxwriter requires a file URI but normalizes relative
                        // file links back to workbook-relative OOXML relationship targets.
                        let url = if target.starts_with("file:///") {
                            target.clone()
                        } else {
                            format!("file:///{}", target.replace('\\', "/"))
                        };
                        let display_text = value.as_str().map(str::to_owned).unwrap_or_else(|| {
                            if value.is_null() { String::new() } else { value.to_string() }
                        });
                        worksheet.write_url_with_options(
                            row as u32, col as u16, url.as_str(), display_text, "", Some(&format),
                        )
                            .map_err(|error| error.to_string())?;
                        continue;
                    }
                }
                write_cell(&mut worksheet, row as u32, col as u16, value, &format)?;
            }
        }
    }
    workbook.save(output_path).map_err(|error| error.to_string())
}

fn column_letters(mut column: u32) -> String {
    let mut result = Vec::new();
    loop {
        result.push((b'A' + (column % 26) as u8) as char);
        if column < 26 { break; }
        column = column / 26 - 1;
    }
    result.iter().rev().collect()
}

fn main() {
    let mut args = env::args_os().skip(1);
    if args.next().as_deref() != Some(std::ffi::OsStr::new("write-tables")) {
        fail("usage: s5-result-writer write-tables <output.xlsx>");
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
