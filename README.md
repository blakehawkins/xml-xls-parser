I encountered some XLS files that fail to be parsed by a number of tools (`xlrd`, `pandas`, `openpyxl`, `calamine`).

The files appear to be in XML format with the following properties:

- `Workbook`
- `Worksheet`
- `Table`
- `Row`
- `Cell`
- `Data`
- `Styles`
- `Style`
- `NumberFormat`
- `Font`
- `Alignment`

It is unclear what makes the files unreadable by XLS and XLSX parsers.

This project reads XLS consisting only of the above properties (XML formatted document) and emits a best-effort TSV.

```bash
$ cp /path/to/file.xls input.xls
$ cargo run > out.tsv
$ less -S out.tsv
```


# How?

It's just a serde specification, using [`serde-xml-rs`](https://lib.rs/crates/serde-xml-rs).

Expect to modify the code if your source document contains anything other than the properties defined above.
