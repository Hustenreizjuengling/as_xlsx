## Usage examples

### Formatting, comments, hyperlinks, merged cells

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell( 5, 1, 5 );
  pck_as_xlsx.cell( 3, 1, 3 );
  pck_as_xlsx.cell( 2, 2, 45 );
  pck_as_xlsx.cell( 3, 2, 'Anton Scheffer', p_alignment => pck_as_xlsx.get_alignment( p_wraptext => true ) );
  pck_as_xlsx.cell( 1, 4, sysdate, p_fontId => pck_as_xlsx.get_font( 'Calibri', p_rgb => 'FFFF0000' ) );
  pck_as_xlsx.cell( 2, 4, sysdate, p_numFmtId => pck_as_xlsx.get_numFmt( 'dd/mm/yyyy h:mm' ) );
  pck_as_xlsx.cell( 3, 4, sysdate, p_numFmtId => pck_as_xlsx.get_numFmt( pck_as_xlsx.orafmt2excel( 'dd/mon/yyyy' ) ) );
  pck_as_xlsx.cell( 5, 5, 75, p_borderId => pck_as_xlsx.get_border( 'double', 'double', 'double', 'double' ) );
  pck_as_xlsx.cell( 2, 3, 33 );
  pck_as_xlsx.hyperlink( 1, 6, 'http://www.amis.nl', 'Amis site' );
  pck_as_xlsx.cell( 1, 7, 'Some merged cells', p_alignment => pck_as_xlsx.get_alignment( p_horizontal => 'center' ) );
  pck_as_xlsx.mergecells( 1, 7, 3, 7 );
  for i in 1 .. 5
  loop
    pck_as_xlsx.comment( 3, i + 3, 'Row ' || (i+3), 'Anton' );
  end loop;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.set_row( 1, p_fillId => pck_as_xlsx.get_fill( 'solid', 'FFFF0000' ) ) ;
  for i in 1 .. 5
  loop
    pck_as_xlsx.cell( 1, i, i );
    pck_as_xlsx.cell( 2, i, i * 3 );
    pck_as_xlsx.cell( 3, i, 'x ' || i * 3 );
  end loop;
  pck_as_xlsx.query2sheet( 'select rownum, x.*
, case when mod( rownum, 2 ) = 0 then rownum * 3 end demo
, case when mod( rownum, 2 ) = 1 then ''demo '' || rownum end demo2 from dual x connect by rownum <= 5' );
  l_blob := pck_as_xlsx.finish;
end;
```

### Data validation (dropdown lists)

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell( 1, 6, 5 );
  pck_as_xlsx.cell( 1, 7, 3 );
  pck_as_xlsx.cell( 1, 8, 7 );
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell( 2, 6, 15, p_sheet => 2 );
  pck_as_xlsx.cell( 2, 7, 13, p_sheet => 2 );
  pck_as_xlsx.cell( 2, 8, 17, p_sheet => 2 );
  pck_as_xlsx.list_validation( 6, 3, 1, 6, 1, 8, p_show_error => true, p_sheet => 1 );
  pck_as_xlsx.defined_name( 2, 6, 2, 8, 'Anton', 2 );
  pck_as_xlsx.list_validation
    ( 6, 1, 'Anton'
    , p_style => 'information'
    , p_title => 'valid values are'
    , p_prompt => '13, 15 and 17'
    , p_show_error => true
    , p_error_title => 'Are you sure?'
    , p_error_txt => 'Valid values are: 13, 15 and 17'
    , p_sheet => 1 );
  l_blob := pck_as_xlsx.finish;
end;
```

### Autofilter

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell( 1, 6, 5 );
  pck_as_xlsx.cell( 1, 7, 3 );
  pck_as_xlsx.cell( 1, 8, 7 );
  pck_as_xlsx.set_autofilter( 1,1, p_row_start => 5, p_row_end => 8 );
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell( 2, 6, 5 );
  pck_as_xlsx.cell( 2, 7, 3 );
  pck_as_xlsx.cell( 2, 8, 7 );
  pck_as_xlsx.set_autofilter( 2,2, p_row_start => 5, p_row_end => 8 );
  l_blob := pck_as_xlsx.finish;
end;
```

### Freeze panes

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.setUseXf( false );
  for c in 1 .. 10
  loop
    pck_as_xlsx.cell( c, 1, 'COL' || c );
    pck_as_xlsx.cell( c, 2, 'val' || c );
    pck_as_xlsx.cell( c, 3, c );
  end loop;
  pck_as_xlsx.freeze_rows( 1 );
  pck_as_xlsx.new_sheet;
  for r in 1 .. 10
  loop
    pck_as_xlsx.cell( 1, r, 'ROW' || r );
    pck_as_xlsx.cell( 2, r, 'val' || r );
    pck_as_xlsx.cell( 3, r, r );
  end loop;
  pck_as_xlsx.freeze_cols( 3 );
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell( 3, 3, 'Start freeze' );
  pck_as_xlsx.freeze_pane( 3,3 );
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
  pck_as_xlsx.add_image( 1, 1, as_barcode.barcode( 'https://github.com/antonscheffer/pck_as_xlsx', 'QR' ) );
  pck_as_xlsx.cell( 1, 8, 'now with png images' );
  l_blob := pck_as_xlsx.finish;
end;
```

### query2sheet with SYS_REFCURSOR and encryption

```sql
declare
  l_blob  blob;
  l_cnt   pls_integer;
  l_query sys_refcursor;
begin
  open l_query for
    select date '1900-02-26' + level "Secret Date"
         , to_char( date '1900-02-26' + level, 'yyyy mon dd' ) "Secret String"
    from dual
    connect by level < 8;
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  l_cnt := pck_as_xlsx.query2sheet
             ( p_rc         => l_query
             , p_sheet      => 1
             , p_col        => 5
             , p_row        => 3
             , p_autofilter => true
             , p_date_format => 'yyyy-mmm-dd'
             , p_title      => 'My Secrets'
             , p_title_xfid => pck_as_xlsx.get_xfid( p_alignment => pck_as_xlsx.get_alignment( p_horizontal => 'centerContinuous' ) )
             );
  pck_as_xlsx.set_column_width( p_col   => 5
                          , p_width => 15
                           );
  pck_as_xlsx.set_column_width( p_col   => 6
                          , p_width => 15
                          );
  pck_as_xlsx.cell( 5
              , l_cnt
                 + 3  -- query start row
                 + 2  -- title + headers
                 + 1  -- interval
              , 'Rows returned: ' || l_cnt );
  -- make sure you have set pck_as_xlsx.use_dbms_crypto = true; in the package specification
  l_blob := pck_as_xlsx.finish( 'demo' );
end;
```

### Formulas

```sql
declare
  l_blob blob;
begin
  pck_as_xlsx.clear_workbook;
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell( 1, 1, 3 );
  pck_as_xlsx.cell( 1, 2, 5 );
  pck_as_xlsx.cell( 1, 3, 4 );
  pck_as_xlsx.num_formula( 2, 1, 'A6+B6' );
  pck_as_xlsx.num_formula( 2, 6, 'SUM(A1:A5)' );
  pck_as_xlsx.num_formula( 1, 6, 'SUM(A1:A2)' );
  pck_as_xlsx.date_formula( 5, 1, 'TODAY()', p_numFmtId => pck_as_xlsx.get_numFmt( 'yyyy-mm-dd' ) );
  pck_as_xlsx.str_formula( 5, 3, 'LOWER(TEXT(TODAY(),"DDDD"))' );
  pck_as_xlsx.new_sheet;
  pck_as_xlsx.cell( 1, 1, 13 );
  pck_as_xlsx.cell( 1, 2, 15 );
  pck_as_xlsx.cell( 1, 3, 14 );
  pck_as_xlsx.num_formula( 2, 1, 'A6+B6' );
  pck_as_xlsx.num_formula( 2, 6, 'SUM(A1:A5)' );
  pck_as_xlsx.num_formula( 1, 6, 'SUM(A1:A2)' );
  l_blob := pck_as_xlsx.finish;
end;
```
