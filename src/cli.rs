use calamine::{open_workbook, DataType, Error, Range, Reader, Xlsx};

fn open_sheet(path: String, sheet_name: String) -> Result<Range<DataType>, Error> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let mut failure_msg = "Cannot find ".to_owned();
    failure_msg.push_str(&sheet_name);
    let range = workbook
        .worksheet_range(&sheet_name)
        .ok_or(Error::Msg("Can't find sheet"))??;
    Ok(range)
}

fn main() {
    println!("Hello from the cli!");

    let path = format!("{}/tests/zipcodes.xlsx", env!("CARGO_MANIFEST_DIR"));
    let sheet_name = "ZIP codes 2018";

    println!("{}", path);

    match open_sheet(path, sheet_name.to_string()) {
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
