use std::io::Write;

use anyhow::{Context, Result};
use serde::Deserialize;

#[derive(Debug, Deserialize)]
struct Alignment {
    #[serde(rename = "Horizontal")]
    horizontal: Option<String>,

    #[serde(rename = "ss:Horizontal")]
    sshorizontal: Option<String>,

    #[serde(rename = "ss:WrapText")]
    sswrap_text: Option<String>,
}

#[derive(Debug, Deserialize)]
struct Font {
    #[serde(rename = "ss:Italic")]
    ssitalic: Option<String>,

    #[serde(rename = "ss:Bold")]
    ssbold: Option<String>,

    #[serde(rename = "ss:Color")]
    sscolor: Option<String>,

    #[serde(rename = "ss:Underline")]
    ssunderline: Option<String>,
}

#[derive(Debug, Deserialize)]
struct NumberFormat {
    #[serde(rename = "ss:Format")]
    ssformat: Option<String>,
}

#[derive(Debug, Deserialize)]
struct Style {
    #[serde(rename = "ss:ID", default)]
    ssid: String,

    #[serde(rename = "Alignment", default)]
    alignments: Vec<Alignment>,

    #[serde(rename = "Font", default)]
    fonts: Vec<Font>,

    #[serde(rename = "NumberFormat", default)]
    number_formats: Vec<NumberFormat>,
}

#[derive(Debug, Deserialize)]
struct Styles {
    #[serde(rename = "Style")]
    styles: Vec<Style>,
}

#[derive(Debug, Deserialize)]
struct Data {
    #[serde(rename = "ss:Type", default)]
    sstype: String,

    #[serde(rename = "$value", default)]
    body: String,
}

#[derive(Debug, Deserialize)]
struct Cell {
    #[serde(rename = "ss:StyleID", default)]
    ssstyleid: String,

    #[serde(rename = "ss:MergeDown", default)]
    ssmerge_down: String,

    #[serde(rename = "Data")]
    data: Vec<Data>,
}

#[derive(Debug, Deserialize)]
struct Row {
    #[serde(rename = "Cell")]
    cells: Vec<Cell>,
}

#[derive(Debug, Deserialize)]
struct Table {
    #[serde(rename = "Row")]
    rows: Vec<Row>,
}

#[derive(Debug, Deserialize)]
struct Worksheet {
    #[serde(rename = "ss:Name", default)]
    ssname: String,

    #[serde(rename = "Table")]
    tables: Vec<Table>,
}

#[derive(Debug, Deserialize)]
struct Workbook {
    #[serde(rename = "xmlns", default)]
    xmlns: String,

    #[serde(rename = "xmlns:ss", default)]
    xmlnsss: String,

    #[serde(rename = "Styles")]
    styles: Vec<Styles>,

    #[serde(rename = "Worksheet")]
    worksheet: Vec<Worksheet>,
}

fn main() -> Result<()> {
    let xd = &mut serde_xml_rs::Deserializer::new_from_reader(std::fs::File::open("input.xls")?);
    let workbook: Workbook = serde_path_to_error::deserialize(xd)?;

    workbook.worksheet.into_iter().for_each(|sheet| {
        sheet.tables.into_iter().for_each(|table| {
            table.rows.into_iter().for_each(|row| {
                let row_fmt = row
                    .cells
                    .into_iter()
                    .map(|cell| {
                        cell.data
                            .into_iter()
                            .map(|d| d.body.replace('\n', " "))
                            .collect::<Vec<_>>()
                            .join("+")
                    })
                    .collect::<Vec<_>>()
                    .join("\t");

                println!("{}", row_fmt);
            })
        })
    });

    // println!("{:?}", workbook);
    std::io::stdout().flush().context("failed to flush stdout")
}
