use chrono::{NaiveDate, NaiveDateTime, Timelike};
use rust_xlsxwriter::{Color, Format, Url, Workbook, Worksheet};
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
    name: String,
    #[serde(default)]
    headers: Vec<Value>,
    #[serde(default)]
    rows: Vec<Vec<Value>>,
    #[serde(default)]
    widths: Vec<f64>,
    #[serde(default)]
    freeze_top_rows: u32,
    #[serde(default)]
    auto_filter: bool,
}

#[derive(Debug, Deserialize)]
struct AuditSummaryPayload {
    sheet_name: String,
    headers: Vec<String>,
    rows: Vec<Vec<String>>,
    #[serde(default)]
    hyperlinks: Vec<AuditHyperlink>,
}

#[derive(Debug, Deserialize)]
struct AuditHyperlink {
    row: u32,
    column: u16,
    target: String,
    location: String,
    display: String,
}

fn fail(message: impl AsRef<str>) -> ! {
    eprintln!("{}", message.as_ref());
    std::process::exit(2);
}

fn safe_number(value: f64) -> f64 {
    format!("{value:.15e}").parse::<f64>().unwrap_or(value)
}

fn excel_serial(date: NaiveDate, seconds: f64) -> f64 {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30).expect("valid Excel epoch");
    (date - epoch).num_days() as f64 + seconds / 86_400.0
}

fn date_value(value: &Value) -> Option<(f64, &'static str)> {
    let object = value.as_object()?;
    let kind = object.get("$audit_datetime")?.as_str()?;
    let text = object.get("value")?.as_str()?;
    if kind == "date" {
        let date = NaiveDate::parse_from_str(text, "%Y-%m-%d").ok()?;
        return Some((excel_serial(date, 0.0), "yyyy-mm-dd"));
    }
    if kind == "datetime" {
        let parsed = NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S%.f")
            .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S"))
            .ok()?;
        let seconds = parsed.num_seconds_from_midnight() as f64
            + f64::from(parsed.and_utc().timestamp_subsec_nanos()) / 1_000_000_000.0;
        return Some((excel_serial(parsed.date(), seconds), "yyyy-mm-dd h:mm:ss"));
    }
    None
}

fn write_value(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Value,
) -> Result<(), String> {
    if value.is_null() {
        worksheet.write_blank(row, col, &Format::new()).map_err(|error| error.to_string())?;
    } else if let Some(text) = value.as_str() {
        worksheet.write_string(row, col, text).map_err(|error| error.to_string())?;
    } else if let Some(flag) = value.as_bool() {
        worksheet.write_boolean(row, col, flag).map_err(|error| error.to_string())?;
    } else if let Some(number) = value.as_f64() {
        worksheet.write_number(row, col, safe_number(number)).map_err(|error| error.to_string())?;
    } else if let Some((number, number_format)) = date_value(value) {
        let format = Format::new().set_num_format(number_format);
        worksheet.write_number_with_format(row, col, number, &format)
            .map_err(|error| error.to_string())?;
    } else {
        return Err(format!("unsupported value at row {}, column {}", row + 1, col + 1));
    }
    Ok(())
}

fn write_table(workbook: &mut Workbook, table: Table) -> Result<(), String> {
    if table.name.chars().count() > 31 {
        return Err(format!("worksheet name exceeds Excel limit: {}", table.name));
    }
    let columns = table.headers.len().max(
        table.rows.iter().map(Vec::len).max().unwrap_or(0),
    );
    if columns > 16_384 {
        return Err(format!("worksheet {} exceeds Excel column limit", table.name));
    }
    let total_rows = table.rows.len() + table.headers.len().min(1);
    if total_rows > 1_048_576 {
        return Err(format!("worksheet {} exceeds Excel row limit", table.name));
    }
    let worksheet = workbook.add_worksheet();
    worksheet.set_name(&table.name).map_err(|error| error.to_string())?;
    if table.freeze_top_rows > 0 {
        worksheet.set_freeze_panes(table.freeze_top_rows, 0)
            .map_err(|error| error.to_string())?;
    }
    if table.auto_filter && columns > 0 {
        let last_row = total_rows.saturating_sub(1) as u32;
        worksheet.autofilter(0, 0, last_row, (columns - 1) as u16)
            .map_err(|error| error.to_string())?;
    }
    for (column, width) in table.widths.iter().enumerate() {
        // Match the raw OOXML width emitted by OpenPyXL. rust_xlsxwriter's
        // character-width conversion adds an implementation margin otherwise.
        worksheet.set_column_width(column as u16, (*width - 0.7109375).max(0.0))
            .map_err(|error| error.to_string())?;
    }
    for (column, value) in table.headers.iter().enumerate() {
        write_value(worksheet, 0, column as u16, value)?;
    }
    for (row_index, row) in table.rows.iter().enumerate() {
        for column in 0..columns {
            let value = row.get(column).unwrap_or(&Value::Null);
            write_value(worksheet, (row_index + 1) as u32, column as u16, value)?;
        }
    }
    Ok(())
}

fn write_payload(payload: Payload, output_path: &Path) -> Result<(), String> {
    let mut workbook = Workbook::new();
    workbook.set_zip_compression_level(6);
    if payload.sheets.is_empty() {
        write_table(&mut workbook, Table {
            name: "汇总".to_string(), headers: Vec::new(), rows: Vec::new(),
            widths: Vec::new(), freeze_top_rows: 0, auto_filter: false,
        })?;
    } else {
        for table in payload.sheets {
            write_table(&mut workbook, table)?;
        }
    }
    workbook.save(output_path).map_err(|error| error.to_string())
}

fn absolute_file_url(path: &str, location: &str) -> String {
    let normalized = path.replace('\\', "/");
    let prefix = if normalized.starts_with("//") {
        "file:"
    } else if normalized.starts_with('/') {
        "file://"
    } else {
        "file:///"
    };
    format!("{prefix}{normalized}#{location}")
}

fn write_audit_summary(payload: AuditSummaryPayload, output_path: &Path) -> Result<(), String> {
    if payload.headers.len() > 16_384 {
        return Err("审核结果表超过 Excel 列数限制".to_string());
    }
    if payload.rows.len() + 1 > 1_048_576 {
        return Err("审核结果表超过 Excel 行数限制".to_string());
    }
    if payload.sheet_name.chars().count() > 31 {
        return Err("审核结果工作表名超过 Excel 31 字符限制".to_string());
    }

    let mut workbook = Workbook::new();
    workbook.set_zip_compression_level(6);
    let worksheet = workbook.add_worksheet();
    worksheet.set_name(&payload.sheet_name).map_err(|error| error.to_string())?;
    worksheet.set_freeze_panes(1, 0).map_err(|error| error.to_string())?;
    if !payload.headers.is_empty() {
        worksheet.autofilter(
            0, 0, payload.rows.len() as u32, (payload.headers.len() - 1) as u16,
        ).map_err(|error| error.to_string())?;
    }

    let header_format = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0xD9EAD3));
    let text_format = Format::new().set_num_format("@");
    let hyperlink_format = Format::new()
        .set_font_size(12.0)
        .set_font_color(Color::RGB(0x0000FF));

    for (column, header) in payload.headers.iter().enumerate() {
        worksheet.write_string_with_format(0, column as u16, header, &header_format)
            .map_err(|error| error.to_string())?;
    }

    let mut links = std::collections::HashMap::new();
    for link in payload.hyperlinks {
        if link.row == 0 || link.row as usize > payload.rows.len()
            || link.column as usize >= payload.headers.len()
        {
            return Err(format!("审核结果超链接坐标越界：row={}, column={}", link.row, link.column));
        }
        links.insert((link.row, link.column), link);
    }

    for (row_index, row) in payload.rows.iter().enumerate() {
        for column in 0..payload.headers.len() {
            let value = row.get(column).map(String::as_str).unwrap_or("");
            let sheet_row = (row_index + 1) as u32;
            if let Some(link) = links.get(&(sheet_row, column as u16)) {
                let url = Url::new(absolute_file_url(&link.target, &link.location))
                    .set_text(link.display.clone());
                worksheet.write_url_with_format(sheet_row, column as u16, url, &hyperlink_format)
                    .map_err(|error| error.to_string())?;
            } else {
                worksheet.write_string_with_format(sheet_row, column as u16, value, &text_format)
                    .map_err(|error| error.to_string())?;
            }
        }
    }

    let widths: [f64; 13] = [38.0, 24.0, 14.0, 16.0, 20.0, 28.0, 42.0, 18.0, 18.0, 18.0, 38.0, 34.0, 24.0];
    for (column, width) in widths.iter().take(payload.headers.len()).enumerate() {
        worksheet.set_column_width(column as u16, (width - 0.7109375).max(0.0))
            .map_err(|error| error.to_string())?;
    }
    workbook.save(output_path).map_err(|error| error.to_string())
}

fn main() {
    let mut args = env::args_os().skip(1);
    let command = args.next();
    if command.as_deref() == Some(std::ffi::OsStr::new("--version")) {
        println!("s1-summary-writer {}; rust_xlsxwriter 0.100.0", env!("CARGO_PKG_VERSION"));
        return;
    }
    let Some(command_name) = command else {
        fail("usage: s1-summary-writer <write-tables|write-audit-summary> <output.xlsx>");
    };
    if command_name != "write-tables" && command_name != "write-audit-summary" {
        fail("usage: s1-summary-writer <write-tables|write-audit-summary> <output.xlsx>");
    }
    let output = match args.next() {
        Some(value) => value,
        None => fail("missing output path"),
    };
    let mut input = String::new();
    if let Err(error) = io::stdin().read_to_string(&mut input) {
        fail(format!("cannot read payload: {error}"));
    }
    let started = Instant::now();
    let result = if command_name == "write-tables" {
        serde_json::from_str::<Payload>(&input)
            .map_err(|error| format!("invalid payload: {error}"))
            .and_then(|payload| write_payload(payload, Path::new(&output)))
    } else {
        serde_json::from_str::<AuditSummaryPayload>(&input)
            .map_err(|error| format!("invalid payload: {error}"))
            .and_then(|payload| write_audit_summary(payload, Path::new(&output)))
    };
    if let Err(error) = result {
        fail(error);
    }
    eprintln!("engine_seconds={:.9}", started.elapsed().as_secs_f64());
}
