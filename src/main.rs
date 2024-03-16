use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

use calamine::{open_workbook_auto, Data, Range, Reader};
use chrono::{NaiveDate, NaiveDateTime};
use clap::{App, Arg};

fn main() {
    let matches = App::new(env!("CARGO_PKG_NAME"))
        .version(env!("CARGO_PKG_VERSION"))
        .author(env!("CARGO_PKG_AUTHORS"))
        .about(env!("CARGO_PKG_DESCRIPTION"))
        .arg(
            Arg::new("file")
                .required(true)
                .help("Excel file to convert to csv"),
        )
        .get_matches();

    println!(
        "{} v{} by {}",
        env!("CARGO_PKG_NAME"),
        env!("CARGO_PKG_VERSION"),
        env!("CARGO_PKG_AUTHORS")
    );

    let file = matches.value_of("file").unwrap();
    let path = PathBuf::from(file);

    // path extension to lowercase
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap()
        .to_lowercase();

    match ext.as_str() {
        "xlsx" | "xlsm" | "xlsb" | "xls" => (),
        _ => panic!("Expecting an Excel file [xlsx, xlsm, xlsb, xls]"),
    }

    let options = sanitize_filename::Options {
        truncate: true,  // truncates to 255 bytes
        windows: true,   // removes reserved names like `con` from start of strings on Windows
        replacement: "", // str to replace sanitized chars
    };

    // get the name of the sheet
    let sheets = open_workbook_auto(&path).unwrap().sheet_names().to_owned();
    let mut excel = open_workbook_auto(&path).unwrap();

    // iterate over all sheets
    for sheet in &sheets {
        if sheet == "hiddenSheet" {
            continue;
        }
        let range = excel.worksheet_range(&sheet).unwrap();

        let extension = ".csv";
        let filename = path.file_stem().unwrap().to_str().unwrap();
        let sheetname = sheet.to_string();
        let sanitized = sanitize_filename::sanitize_with_options(
            format!("{}_{}{}", filename, sheetname, extension),
            options.clone(),
        );

        let dest = path.with_file_name(sanitized);
        println!("{}", dest.display());
        let mut dest = BufWriter::new(File::create(dest).unwrap());

        write(&mut dest, &range).unwrap();
    }
}

fn write<W: Write>(dest: &mut W, range: &Range<Data>) -> std::io::Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            match *c {
                Data::Empty => Ok(()),
                Data::String(ref s) | Data::DateTimeIso(ref s) | Data::DurationIso(ref s) => {
                    write!(dest, "{}", s)
                }
                Data::Float(ref f) => write!(dest, "{}", f),
                // Datetime as YYYY-MM-DDTHH:MM:SS
                Data::DateTime(ref d) => write!(dest, "{}", convert_excel_date_time(d.as_f64())),
                Data::Int(ref i) => write!(dest, "{}", i),
                Data::Error(ref e) => write!(dest, "{:?}", e),
                Data::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != n {
                write!(dest, ";")?;
            }
        }
        write!(dest, "\r\n")?;
    }
    Ok(())
}

#[allow(deprecated)]
fn convert_excel_date_time(excel_datetime: f64) -> NaiveDateTime {
    let days = excel_datetime.floor() as i32;
    let seconds = ((excel_datetime - days as f64) * 86_400.0) as u32;

    let base_date = NaiveDate::from_ymd(1899, 12, 30);
    base_date.and_hms(0, 0, 0)
        + chrono::Duration::days(days as i64)
        + chrono::Duration::seconds(seconds as i64)
}
