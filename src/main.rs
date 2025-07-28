use actix_web::{post, web, App, HttpResponse, HttpServer, Responder};
use serde::Deserialize;
use std::io::{Cursor};

// Excel
use rust_xlsxwriter::Workbook;
use std::fs;
use tempfile::NamedTempFile;
// CSV
use csv;

use docx_rs::*;

// PDF (printpdf 0.3.1)

#[derive(Debug, Deserialize)]
pub struct TableData {
    pub headers: Vec<String>,
    pub rows: Vec<Vec<String>>,
    pub title: Option<String>,
}

fn export_csv(data: &TableData) -> HttpResponse {
    let mut wtr = csv::Writer::from_writer(vec![]);
    let _ = wtr.write_record(&data.headers);
    for row in &data.rows {
        let _ = wtr.write_record(row);
    }
    let bytes = wtr.into_inner().unwrap();
    HttpResponse::Ok()
        .append_header(("Content-Type", "text/csv"))
        .append_header(("Content-Disposition", "attachment; filename=data.csv"))
        .body(bytes)
}

fn export_excel(data: &TableData) -> HttpResponse {
    // 1) Crée un nouveau classeur
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // 2) Titre optionnel en A1
    if let Some(title) = &data.title {
        worksheet.write(0, 0, title).unwrap();
    }

    // 3) En‑têtes en ligne 1
    for (col, header) in data.headers.iter().enumerate() {
        worksheet.write(1, col as u16, header).unwrap();
    }

    // 4) Données à partir de la ligne 2
    for (row_idx, row) in data.rows.iter().enumerate() {
        for (col_idx, val) in row.iter().enumerate() {
            worksheet.write((row_idx + 2) as u32, col_idx as u16, val).unwrap();
        }
    }

    // 5) Sauvegarde dans un fichier temporaire
    let tmp = NamedTempFile::new().unwrap();
    let path = tmp.path().to_str().unwrap();
    workbook.save(path).unwrap();

    // 6) Lit le fichier en mémoire
    let bytes = fs::read(path).unwrap();

    // 7) Renvoie la réponse
    HttpResponse::Ok()
        .append_header((
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ))
        .append_header(("Content-Disposition", "attachment; filename=data.xlsx"))
        .body(bytes)
}


fn export_docx(data: &TableData) -> HttpResponse {
    let mut doc = Docx::new();
    if let Some(title) = &data.title {
        doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_text(title)));
    }
    let mut rows: Vec<TableRow> = Vec::new();

    // Header
    let header_cells: Vec<TableCell> = data
        .headers
        .iter()
        .map(|h| TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(h))))
        .collect();
    rows.push(TableRow::new(header_cells));

    // Body
    for row in &data.rows {
        let cells: Vec<TableCell> = row
            .iter()
            .map(|cell| {
                TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(cell)))
            })
            .collect();
        rows.push(TableRow::new(cells));
    }

    doc = doc.add_table(Table::new(rows));

    let mut buf = Cursor::new(Vec::new());
    doc.build().pack(&mut buf).unwrap();
    let bytes = buf.into_inner();
    HttpResponse::Ok()
        .append_header((
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ))
        .append_header(("Content-Disposition", "attachment; filename=data.docx"))
        .body(bytes)
}


#[post("/export/{format}")]
async fn export_file(
    format: web::Path<String>,
    data: web::Json<TableData>,
) -> impl Responder {
    match format.as_str() {
        "csv" => export_csv(&data),
        "xlsx" => export_excel(&data),
        "docx" => export_docx(&data),
        _ => HttpResponse::BadRequest().body("Format non supporté"),
    }
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    HttpServer::new(|| App::new().service(export_file))
        .bind(("127.0.0.1", 8781))?
        .run()
        .await
}
