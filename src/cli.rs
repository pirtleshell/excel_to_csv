use calamine::{open_workbook, DataType, Error, Range, Reader, Xlsx};
use std::env;

fn main() {
    let args: Vec<String> = env::args().collect();
    match args.len() {
        s if s < 2 => usage("Missing filename"),
        s if s > 2 => usage(&format!("Found more args than expected: {:?}", &args[1..])),
        _ => (),
    };

    let filename = &args[1];

    if filename == "--help" || filename == "-h" {
        usage("");
    }

    let sheet_name = "ZIP codes 2018";

    match open_sheet(filename.to_string(), sheet_name.to_string()) {
        Ok(sheet) => {
            println!("nice!");

            let rows = sheet.rows();
            let headers = match sheet.rows().next() {
                Some(h) => h,
                None => panic!("No data in sheet found."),
            };

            for row in rows.take(5) {
                println!("\n#####\n");

                let r = row.into_iter().zip(headers.into_iter());
                for (data, header) in r {
                    println!("{}: {}", header, data);
                }

                println!("\n#####\n");
            }
        }
        Err(m) => println!("booo\n{}", m),
    }
}

fn usage(msg: &str) {
    if !msg.is_empty() {
        println!("{}", msg);
    }
    println!("Usage: excel_to_csv ./path/to/spreadsheet.xlsx");
    std::process::exit(0x0100);
}

fn open_sheet(path: String, sheet_name: String) -> Result<Range<DataType>, Error> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let mut failure_msg = "Cannot find ".to_owned();
    failure_msg.push_str(&sheet_name);
    let range = workbook
        .worksheet_range(&sheet_name)
        .ok_or(Error::Msg("Can't find sheet"))??;
    Ok(range)
}
