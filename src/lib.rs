use rust_xlsxwriter::{Workbook, XlsxError};
use serde_json::Value;
use std::io::BufRead;

#[derive(Debug)]
pub enum Error {
    Io(std::io::Error),
    Json(serde_json::Error),
    Xlsx(XlsxError),
}

impl From<std::io::Error> for Error {
    fn from(err: std::io::Error) -> Self {
        Error::Io(err)
    }
}

impl From<serde_json::Error> for Error {
    fn from(err: serde_json::Error) -> Self {
        Error::Json(err)
    }
}

impl From<XlsxError> for Error {
    fn from(err: XlsxError) -> Self {
        Error::Xlsx(err)
    }
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match self {
            Error::Io(err) => write!(f, "IO error: {}", err),
            Error::Json(err) => write!(f, "JSON error: {}", err),
            Error::Xlsx(err) => write!(f, "XLSX error: {}", err),
        }
    }
}

pub fn run<R: BufRead>(reader: R, output: &str, sheet: &str, row_limit: u32) -> Result<(), Error> {
    let mut json_data: Vec<Value> = Vec::new();

    for line in reader.lines().take(row_limit as usize) {
        let line = line?;
        let value: Value = serde_json::from_str(&line)?;
        json_data.push(value);
    }

    if json_data.is_empty() {
        println!("No JSON data to write to Excel.");
        return Ok(());
    }

    let mut headers: Vec<String> = Vec::new();
    if let Some(obj) = json_data.first().and_then(|v| v.as_object()) {
        for key in obj.keys() {
            headers.push(key.clone());
        }
    }

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet().set_name(sheet)?;

    // Write headers
    for (col_num, header) in headers.iter().enumerate() {
        worksheet.write_string(0, col_num as u16, header)?;
    }

    // Write data
    for (row_num, json_obj) in json_data.iter().enumerate() {
        if let Some(obj) = json_obj.as_object() {
            for (col_num, header) in headers.iter().enumerate() {
                if let Some(value) = obj.get(header) {
                    match value {
                        Value::String(s) => {
                            worksheet.write_string((row_num + 1) as u32, col_num as u16, s)?;
                        }
                        Value::Number(n) => {
                            if let Some(f) = n.as_f64() {
                                worksheet.write_number((row_num + 1) as u32, col_num as u16, f)?;
                            } else if let Some(i) = n.as_i64() {
                                worksheet.write_number(
                                    (row_num + 1) as u32,
                                    col_num as u16,
                                    i as f64,
                                )?;
                            } else {
                                worksheet.write_string(
                                    (row_num + 1) as u32,
                                    col_num as u16,
                                    n.to_string(),
                                )?;
                            }
                        }
                        Value::Bool(b) => {
                            worksheet.write_boolean((row_num + 1) as u32, col_num as u16, *b)?;
                        }
                        Value::Null => {
                            // Write empty string or leave blank
                            worksheet.write_string((row_num + 1) as u32, col_num as u16, "")?;
                        }
                        _ => {
                            // For Array, Object, etc., write their string representation
                            worksheet.write_string(
                                (row_num + 1) as u32,
                                col_num as u16,
                                value.to_string(),
                            )?;
                        }
                    }
                }
            }
        }
    }

    workbook.save(output)?;

    Ok(())
}
