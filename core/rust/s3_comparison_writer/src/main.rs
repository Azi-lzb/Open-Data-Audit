use rust_xlsxwriter::{Color, Format, Workbook, Worksheet};
use serde::Deserialize;
use serde_json::Value;
use std::collections::HashMap;
use std::env;
use std::io::{self, Read};
use std::path::Path;
use std::time::Instant;

#[derive(Debug, Deserialize)]
struct Payload {
    title: String,
    headers: Vec<String>,
    rows: Vec<Vec<Value>>,
    #[serde(default)]
    fills: Vec<RowFill>,
}

#[derive(Debug, Deserialize)]
struct RowFill {
    row: usize,
    color: String,
    scope: String,
}

fn fail(message: impl AsRef<str>) -> ! {
    eprintln!("{}", message.as_ref());
    std::process::exit(2);
}

fn safe_number(value: f64) -> f64 {
    // OpenPyXL saves floats using a 15 significant digit safe string.
    format!("{value:.15e}").parse::<f64>().unwrap_or(value)
}

fn parse_rgb(value: &str) -> Result<Color, String> {
    let hex = value.get(value.len().saturating_sub(6)..)
        .ok_or_else(|| format!("invalid fill color: {value}"))?;
    let rgb = u32::from_str_radix(hex, 16)
        .map_err(|_| format!("invalid fill color: {value}"))?;
    Ok(Color::RGB(rgb))
}

fn write_value(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Value,
    format: &Format,
    has_format: bool,
) -> Result<(), String> {
    if value.is_null() {
        if has_format {
            worksheet.write_blank(row, col, format).map_err(|error| error.to_string())?;
        }
    } else if let Some(text) = value.as_str() {
        worksheet.write_string_with_format(row, col, text, format).map_err(|error| error.to_string())?;
    } else if let Some(flag) = value.as_bool() {
        worksheet.write_boolean_with_format(row, col, flag, format).map_err(|error| error.to_string())?;
    } else if let Some(number) = value.as_f64() {
        worksheet.write_number_with_format(row, col, safe_number(number), format).map_err(|error| error.to_string())?;
    } else {
        return Err(format!("unsupported value at row {}, column {}", row + 1, col + 1));
    }
    Ok(())
}

fn fill_map(fills: &[RowFill]) -> HashMap<usize, (&str, &str)> {
    fills.iter().map(|fill| (fill.row, (fill.color.as_str(), fill.scope.as_str()))).collect()
}

fn cell_format(col: usize, fill: Option<&str>) -> Result<(Format, bool), String> {
    let mut format = Format::new();
    let mut has_format = false;
    let number_format = match col {
        1 => Some("yyyy/m/d"),
        13 | 14 | 15 => Some("0.00"),
        16 => Some("0.00%"),
        _ => None,
    };
    if let Some(number_format) = number_format {
        format = format.set_num_format(number_format);
        has_format = true;
    }
    if let Some(color) = fill {
        format = format.set_background_color(parse_rgb(color)?);
        has_format = true;
    }
    Ok((format, has_format))
}

fn write_payload(payload: Payload, output_path: &Path) -> Result<(), String> {
    let mut workbook = Workbook::new();
    workbook.set_zip_compression_level(6);
    let worksheet = workbook.add_worksheet();
    if payload.title.chars().count() > 31 {
        return Err(format!("worksheet name exceeds Excel limit: {}", payload.title));
    }
    worksheet.set_name(&payload.title).map_err(|error| error.to_string())?;
    worksheet.set_freeze_panes(1, 0).map_err(|error| error.to_string())?;
    let last_row = payload.rows.len().max(1) as u32;
    let last_col = payload.headers.len().saturating_sub(1) as u16;
    if !payload.headers.is_empty() {
        worksheet.autofilter(0, 0, last_row, last_col).map_err(|error| error.to_string())?;
    }
    let widths: [f64; 22] = [10.0, 12.0, 10.0, 24.0, 10.0, 10.0, 12.0, 10.0, 32.0, 9.0, 9.0, 6.0, 6.0,
                  14.0, 14.0, 14.0, 12.0, 24.0, 24.0, 24.0, 40.0, 12.0];
    for (column, width) in widths.iter().enumerate() {
        // Match the raw OOXML width used by OpenPyXL, compensating for the
        // character-to-pixel conversion performed by rust_xlsxwriter.
        worksheet.set_column_width(column as u16, (*width - 0.7109375).max(0.0))
            .map_err(|error| error.to_string())?;
    }

    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color(Color::RGB(0x1F4E78))
        .set_text_wrap();
    let mut style_cache: HashMap<String, (Format, bool)> = HashMap::new();
    for (col, value) in payload.headers.iter().enumerate() {
        worksheet.write_string_with_format(0, col as u16, value, &header_format)
            .map_err(|error| error.to_string())?;
    }

    let fills = fill_map(&payload.fills);
    for (row_index, cells) in payload.rows.iter().enumerate() {
        let excel_row = row_index + 1;
        for col in 0..payload.headers.len() {
            let fill = fills.get(&row_index).and_then(|(color, scope)| {
                if (*scope == "row" && col < 18) || (*scope == "remark" && col == 17) {
                    Some(*color)
                } else {
                    None
                }
            });
            let cache_key = format!("{col}|{}", fill.unwrap_or(""));
            if !style_cache.contains_key(&cache_key) {
                style_cache.insert(cache_key.clone(), cell_format(col, fill)?);
            }
            let (format, has_format) = style_cache.get(&cache_key).unwrap();
            let value = cells.get(col).unwrap_or(&Value::Null);
            write_value(worksheet, excel_row as u32, col as u16, value, format, *has_format)?;
        }
    }
    workbook.save(output_path).map_err(|error| error.to_string())
}

fn main() {
    let mut args = env::args_os().skip(1);
    let command = args.next();
    if command.as_deref() == Some(std::ffi::OsStr::new("--version")) {
        println!(
            "s3-comparison-writer {}; rust_xlsxwriter 0.100.0",
            env!("CARGO_PKG_VERSION")
        );
        return;
    }
    if command.as_deref() != Some(std::ffi::OsStr::new("write-comparison")) {
        fail("usage: s3-comparison-writer write-comparison <output.xlsx>");
    }
    let output = match args.next() {
        Some(value) => value,
        None => fail("missing output path"),
    };
    let mut input = String::new();
    if let Err(error) = io::stdin().read_to_string(&mut input) {
        fail(format!("cannot read payload: {error}"));
    }
    let payload: Payload = serde_json::from_str(&input).unwrap_or_else(|error| fail(format!("invalid payload: {error}")));
    let started = Instant::now();
    if let Err(error) = write_payload(payload, Path::new(&output)) {
        fail(error);
    }
    eprintln!("engine_seconds={:.9}", started.elapsed().as_secs_f64());
}
