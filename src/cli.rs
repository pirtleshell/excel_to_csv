use calamine::{open_workbook, DataType, Error, Range, Reader, Xlsx};
use std::env;

fn main() {
    let args: Vec<String> = env::args().collect();

    match args.len() {
        1 => usage("Missing file & sheet name."),
        2 => {
            if &args[1] == "--help" || &args[1] == "-h" {
                usage("");
            }
            let workbook: Option<Xlsx<_>> = match open_workbook(&args[1]) {
                Ok(wb) => Some(wb),
                Err(_) => None,
            };

            if let Some(wb) = workbook {
                println!("Available sheets:");
                for sheet in wb.sheet_names() {
                    println!("* '{}'", sheet);
                }
                println!("----");
            } else {
                println!("Unable to find sheet names in '{}'", &args[1])
            }

            usage("Missing sheet name.")
        }
        s if s > 3 => usage(&format!("Found more args than expected: {:?}", &args[1..])),
        _ => (),
    };

    let filename = &args[1];
    let sheet_name = &args[2];

    match open_sheet(filename.to_string(), sheet_name.to_string()) {
        Ok(sheet) => {
            let rows = sheet.rows().take(100);
            for row in rows {
                let r = row.into_iter();
                let last_idx = r.len() - 1;
                for (col_idx, data) in r.enumerate() {
                    // escape double quotes with double double quotes!
                    let value = data.to_string().replace("\"", "\"\"");

                    if value.len() != 0 {
                        // wrap in double quotes
                        print!("\"{}\"", value);
                    }

                    if col_idx != last_idx {
                        print!(",");
                    }
                }
                println!();
            }
        }
        Err(m) => panic!("booo!\n{}", m),
    }
}

fn usage(msg: &str) {
    if !msg.is_empty() {
        println!("{}", msg);
    }
    println!("Usage: excel_to_csv ./path/to/spreadsheet.xlsx 'Sheet Name'");
    std::process::exit(0x0100);
}

fn open_sheet(path: String, sheet_name: String) -> Result<Range<DataType>, Error> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook
        .worksheet_range(&sheet_name)
        .ok_or(Error::Msg("Can't find sheet"))??;
    Ok(range)
}
