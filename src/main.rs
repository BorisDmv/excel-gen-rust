extern crate simple_excel_writer as excel;
use excel::*;
use mysql::{self, OptsBuilder};
use mysql::prelude::{Queryable};


fn main() {


    create_excel();

}

fn create_excel() {


    let mut wb = Workbook::create("tmp/b.xlsx");

    let mut sheet = wb.create_sheet("Sheet");

    sheet.add_column(Column { width: 20.0 });
    sheet.add_column(Column { width: 20.0 });

    wb.write_sheet(&mut sheet, |sheet_writer| {

        let sw = sheet_writer;

        // Define connection parameters
        let server = "localhost";
        let database = "test-db";
        let username = "root";
        let password = "";
        let port = 3306; // 3306 otherwise 3307 for ssh tunnel

        // Build options for the database connection
        let opts = OptsBuilder::new()
            .ip_or_hostname(Some(server))
            .db_name(Some(database))
            .user(Some(username))
            .pass(Some(password))
            .tcp_port(port)
            .ssl_opts(None);

        // Connect to the database
        let pool = mysql::Pool::new(mysql::Opts::from(opts)).unwrap();
        let mut conn = pool.get_conn().unwrap();

        let result: Vec<(String, String, String, String, String)> = conn
            .query("SELECT id, client, quarter, year, wages FROM taxable_wages")
            .unwrap();

        for db_row in result {
            sw.append_row(row![db_row.0, db_row.1, db_row.2, db_row.3, db_row.4])?;
        }


        sw.append_row(row!["END", "LINE", true])

    }).expect("write excel error!");

    wb.close().expect("close excel error!");
}