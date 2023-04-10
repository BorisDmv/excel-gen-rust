use printpdf::*;
use std::fs::File;
use std::io::BufWriter;
extern crate timer;
extern crate chrono;
use std::sync::mpsc::channel;
#[macro_use]
extern crate simple_excel_writer as excel;
use excel::*;



fn main() {

    let timer = timer::Timer::new();
    let (tx, rx) = channel();
    timer.schedule_with_delay(chrono::Duration::seconds(3), move || {
        tx.send(()).unwrap();
    });

    create_excel();


    rx.recv().unwrap();
    let elapsed = timer.elapsed();
    println!("This code has been executed in {} seconds", elapsed.as_secs());
}

fn create_excel() {
    let mut wb = Workbook::create("tmp/b.xlsx");
    //let mut sheet = wb.create_sheet("SheetName");

    // set column width
    // sheet.add_column(Column { width: 30.0 });
    // sheet.add_column(Column { width: 30.0 });
    // sheet.add_column(Column { width: 80.0 });
    // sheet.add_column(Column { width: 60.0 });

    // wb.write_sheet(&mut sheet, |sheet_writer| {
    //     let sw = sheet_writer;
    //     sw.append_row(row!["Name", "Title","Success","XML Remark"])?;

    //     sw.append_row(row!["Amy", "wtf", true,"<xml><tag>\"Hello\" & 'World'</tag></xml>"])?;
    //     sw.append_row(row!["Tony", blank!(2), "retired"])
    // }).expect("write excel error!");

    let mut sheet = wb.create_sheet("Sheet2");
    wb.write_sheet(&mut sheet, |sheet_writer| {

        let sw = sheet_writer;

        for i in 1..600 {
            sw.append_row(row!["Name", "Title","Success","Remark"])?;
        }
            //sw.append_row(row!["Name", "Title","Success","Remark"])?;
            sw.append_row(row!["Amy", "Manager", true])
        
    }).expect("write excel error!");

    wb.close().expect("close excel error!");
}

fn create_pdf(){
        let (doc, page1, layer1) = PdfDocument::new("PDF_Document_title", Mm(1280.0), Mm(720.0), "Layer 1");
        let current_layer = doc.get_page(page1).get_layer(layer1);


        let text = "Lorem ipsum";
        let font = doc.add_external_font(File::open("assets/fonts/RobotoMedium.ttf").unwrap()).unwrap();

        current_layer.use_text(text, 48.0, Mm(200.0), Mm(200.0), &font);


        // Save the PDF document to a file
        doc.save(&mut BufWriter::new(File::create("test_working.pdf").unwrap())).unwrap();
}