use std::io::Write;
use std::env;
use std::fs::File;
use flate2::read::GzEncoder;
use flate2::Compression;
use std::fmt::format;
use std::process::Command;
use std::env::args;

fn main() {
    let params:Vec<String> = env::args().collect();

    if params.len() != 4 { return ;}
    let from = params.get(1).unwrap();
    let to = params.get(2).unwrap();
    let name = params.get(3).unwrap();

    // Открытие CSV файла
    let reader = csv::Reader::from_path(from);
    if reader.is_err() { return; }
    let mut reader = reader.unwrap();
    // Заголовоки колонок
    let headers = reader.headers();
    if headers.is_err() { return; }
    let headers = headers.unwrap().iter().map(|s|s.to_string()).collect::<Vec<String>>();

    // Разбитие на несколько файлов
    let mut num_file = 1;
    'new_file:loop {
        let mut number_row = 0;
        let name_file = to.to_string() + format!("{}_{}.xls", name, num_file).as_str();
        // xls
        let mut xls = xlsxwriter::Workbook::new(&name_file);
        let sheet = xls.add_worksheet(None);
        if sheet.is_err() {
            xls.close();
            return;
        }
        let mut sheet = sheet.unwrap();
        // Заголовоки
        headers.iter().enumerate().for_each(|(num_col,text)|{
            sheet.write_string(number_row, num_col as u16, text, None);
        });
        // Обработка строк
        for row in reader.records() {
            if row.is_err() { continue; }
            number_row += 1;

            row.as_ref().unwrap().iter().enumerate().for_each(|(num_col,text)|{
                sheet.write_string(number_row, num_col as u16, text, None);
            });

            if number_row >= 65000{
                xls.close();
                num_file += 1;
                continue 'new_file;
            }
        }
        break;
    }

    // Если файлов несколько собираем в единный zip
    if num_file > 1 {
        Command::new("zip")
            .arg("-j")
            .arg("-m")
            .arg(format!("{}/{}.zip", to, name))
            .args({1..=num_file}.map(|num|{
                    format!("{}/{}_{}.xls", to, name, num)
                }))
            .output()
            .expect("error");

        // let tar_gz = File::create(format!("{}/{}",to,name)+".tar.gz").unwrap();
        // let enc = GzEncoder::new(tar_gz, Compression::default());
        // let mut tar = tar::Builder::new(enc);
        // {1..=num_file}.for_each(|num|{
        //     let f = format!("{}/{}_{}.xls", to, name, num).replace("//","/");
        //     let mut fo = File::open(f ).unwrap();
        //     tar.append_file(
        //         format!("./{}_{}.xls",name, num),
        //         &mut fo
        //     );
        // });
        // tar.finish();
    }else{
        std::fs::rename(
            format!("{}/{}_1.xls", to, name),
            format!("{}/{}.xls", to, name)
        );
    }
}
