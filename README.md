# pck_as_xlsx

Oracle PL/SQL package for generating Excel `.xlsx` files directly from the database, returned as `BLOB`.

> Fork of [as_xlsx](https://technology.amis.nl/languages/oracle-plsql/create-an-excel-file-with-plsql/) by Anton Scheffer.
> Original includes read functionality and UTL_FILE support. This fork is refactored to a **write-only** package — all read/file-I/O code has been removed and the internal `finish()` function has been modularized.

## Features

- Create multi-sheet Excel workbooks
- Write cell values: numbers, strings, dates
- Cell formulas (numeric, string, date)
- Fonts, fills, borders, alignment, number formats
- Merged cells, comments, hyperlinks
- Column/row formatting, column widths, row heights
- Freeze panes, autofilters, table styles
- Data validation (dropdown lists)
- Named ranges
- Embedded images (PNG, JPG, GIF, BMP)
- Query-to-sheet: populate sheets directly from SQL or `SYS_REFCURSOR`
- Optional password encryption (requires `DBMS_CRYPTO`)

## Installation

Compile both files on your Oracle database in order:

```sql
@src/pck_as_xlsx.pks
@src/pck_as_xlsx.pkb
```

### Encryption (optional)

To enable password-protected XLSX output, set the constant in `pck_as_xlsx.pks` before compiling:

```sql
use_dbms_crypto constant boolean := true;
```

This requires the `DBMS_CRYPTO` package grant.

## Usage

### Basic example

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet('Demo');
  pck_as_xlsx.cell(1, 1, 'Hello');
  pck_as_xlsx.cell(2, 1, 42);
  pck_as_xlsx.cell(3, 1, sysdate, p_numFmtId => pck_as_xlsx.get_numFmt('dd/mm/yyyy'));
  l_blob := pck_as_xlsx.finish;
  -- use l_blob: store it, return via web, etc.
end;
```

### Formatting

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell(1, 1, 'Bold red',
    p_fontId => pck_as_xlsx.get_font('Calibri', p_bold => true, p_rgb => 'FFFF0000'));
  pck_as_xlsx.cell(1, 2, 'Wrapped text',
    p_alignment => pck_as_xlsx.get_alignment(p_wraptext => true));
  pck_as_xlsx.cell(2, 1, 100,
    p_borderId => pck_as_xlsx.get_border('double', 'double', 'double', 'double'));
  pck_as_xlsx.set_column_width(1, 20);
  l_blob := pck_as_xlsx.finish;
end;
```

### Query to sheet

```sql
declare
  l_blob blob;
  l_cnt  pls_integer;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet('Report');
  l_cnt := pck_as_xlsx.query2sheet(
    p_sql        => 'select employee_id, first_name, hire_date from employees',
    p_autofilter => true,
    p_date_format => 'yyyy-mm-dd'
  );
  l_blob := pck_as_xlsx.finish;
end;
```

### Freeze panes and autofilter

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  for c in 1 .. 10 loop
    pck_as_xlsx.cell(c, 1, 'COL' || c);
    pck_as_xlsx.cell(c, 2, 'val' || c);
    pck_as_xlsx.cell(c, 3, c);
  end loop;
  pck_as_xlsx.freeze_rows(1);
  pck_as_xlsx.set_autofilter(1, 10, 1, 3);
  l_blob := pck_as_xlsx.finish;
end;
```

### Formulas

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell(1, 1, 3);
  pck_as_xlsx.cell(1, 2, 5);
  pck_as_xlsx.cell(1, 3, 4);
  pck_as_xlsx.num_formula(2, 1, 'SUM(A1:A3)');
  pck_as_xlsx.date_formula(3, 1, 'TODAY()', p_numFmtId => pck_as_xlsx.get_numFmt('yyyy-mm-dd'));
  pck_as_xlsx.str_formula(4, 1, 'LOWER(TEXT(TODAY(),"DDDD"))');
  l_blob := pck_as_xlsx.finish;
end;
```

### Data validation

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell(1, 1, 'A');
  pck_as_xlsx.cell(1, 2, 'B');
  pck_as_xlsx.cell(1, 3, 'C');
  -- dropdown at B1 referencing A1:A3
  pck_as_xlsx.list_validation(2, 1, 1, 1, 1, 3, p_show_error => true);
  l_blob := pck_as_xlsx.finish;
end;
```

### Images

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.add_image(1, 1, my_image_blob, p_name => 'logo');
  l_blob := pck_as_xlsx.finish;
end;
```

## API Reference

| Procedure / Function | Description |
|---|---|
| `clear_workbook` | Reset workbook state |
| `new_sheet` | Add a new worksheet |
| `cell` | Write a value (number, varchar2, or date) |
| `hyperlink` | Add a clickable hyperlink |
| `num_formula` / `str_formula` / `date_formula` | Write a formula |
| `comment` | Add a cell comment |
| `mergecells` | Merge a range of cells |
| `list_validation` | Add dropdown validation |
| `defined_name` | Create a named range |
| `set_column_width` | Set column width |
| `set_column` | Set default column formatting |
| `set_row` | Set row formatting / height |
| `freeze_rows` / `freeze_cols` / `freeze_pane` | Freeze panes |
| `set_autofilter` | Enable autofilter on a range |
| `set_table` | Format a range as a table |
| `set_tabcolor` | Set worksheet tab color |
| `add_image` | Embed an image |
| `query2sheet` | Populate sheet from SQL / `SYS_REFCURSOR` |
| `setUseXf` | Toggle XF style mode for `query2sheet` |
| `get_numFmt` / `get_font` / `get_fill` / `get_border` / `get_alignment` / `get_xfid` | Style helpers |
| `OraFmt2Excel` | Convert Oracle date format to Excel format |
| `finish` | Generate and return the XLSX as `BLOB` |
| `get_version` | Return package version string |

## License

MIT License — see [LICENSE](LICENSE).

Original copyright (c) Anton Scheffer.
