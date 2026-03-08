create or replace package body pck_as_xlsx
is
  --------------------------------------------------------------------------
  -- Constants
  --------------------------------------------------------------------------
  c_version constant varchar2(20) := 'pck_as_xlsx60';

  -- ZIP format signatures
  c_lob_duration constant pls_integer := dbms_lob.call;
  c_LOCAL_FILE_HEADER        constant raw(4) := hextoraw( '504B0304' );
  c_CENTRAL_FILE_HEADER      constant raw(4) := hextoraw( '504B0102' );
  c_END_OF_CENTRAL_DIRECTORY constant raw(4) := hextoraw( '504B0506' );

  --------------------------------------------------------------------------
  -- Private type declarations
  --------------------------------------------------------------------------
  type tp_XF_fmt is record
    ( numFmtId pls_integer
    , fontId pls_integer
    , fillId pls_integer
    , borderId pls_integer
    , alignment tp_alignment
    , height number
    );
  type tp_col_fmts is table of tp_XF_fmt index by pls_integer;
  type tp_row_fmts is table of tp_XF_fmt index by pls_integer;
  type tp_widths is table of number index by pls_integer;
  type tp_cell is record
    ( value   number
    , style   varchar2(50)
    , formula varchar2(1024)
    );
  type tp_cells is table of tp_cell index by pls_integer;
  type tp_rows is table of tp_cells index by pls_integer;
  type tp_autofilter is record
    ( column_start pls_integer
    , column_end   pls_integer
    , row_start    pls_integer
    , row_end      pls_integer
    );
  type tp_autofilters is table of tp_autofilter index by pls_integer;
  type tp_table is record
    ( sheet        pls_integer
    , column_start pls_integer
    , column_end   pls_integer
    , row_start    pls_integer
    , row_end      pls_integer
    , style        varchar2(1000)
    , name         varchar2(32767)
    );
  type tp_tables is table of tp_table index by pls_integer;
  type tp_hyperlink is record
    ( cell varchar2(10)
    , url  varchar2(1000)
    , location varchar2(100)
    , tooltip varchar2(1000)
    );
  type tp_hyperlinks is table of tp_hyperlink index by pls_integer;
  subtype tp_author is varchar2(32767 char);
  type tp_authors is table of pls_integer index by tp_author;
  authors tp_authors;
  type tp_comment is record
    ( text varchar2(32767 char)
    , author tp_author
    , row pls_integer
    , column pls_integer
    , width pls_integer
    , height pls_integer
    );
  type tp_comments is table of tp_comment index by pls_integer;
  type tp_mergecells is table of varchar2(21) index by pls_integer;
  type tp_validation is record
    ( type varchar2(10)
    , errorstyle varchar2(32)
    , showinputmessage boolean
    , prompt varchar2(32767 char)
    , title varchar2(32767 char)
    , error_title varchar2(32767 char)
    , error_txt varchar2(32767 char)
    , showerrormessage boolean
    , formula1 varchar2(32767 char)
    , formula2 varchar2(32767 char)
    , allowBlank boolean
    , sqref varchar2(32767 char)
    );
  type tp_validations is table of tp_validation index by pls_integer;
  type tp_drawing is record
    ( img_id pls_integer
    , row pls_integer
    , col pls_integer
    , scale number
    , name varchar2(100)
    , title varchar2(100)
    , description varchar2(4000)
    );
  type tp_drawings is table of tp_drawing index by pls_integer;
  type tp_sheet is record
    ( rows tp_rows
    , widths tp_widths
    , name varchar2(100)
    , freeze_rows pls_integer
    , freeze_cols pls_integer
    , autofilters tp_autofilters
    , hyperlinks tp_hyperlinks
    , col_fmts tp_col_fmts
    , row_fmts tp_row_fmts
    , comments tp_comments
    , mergecells tp_mergecells
    , validations tp_validations
    , drawings tp_drawings
    , tabcolor varchar2(8)
    , show_gridlines boolean
    , grid_color_idx pls_integer
    , show_headers   boolean
    );
  type tp_sheets is table of tp_sheet index by pls_integer;
  type tp_numFmt is record
    ( numFmtId pls_integer
    , formatCode varchar2(100)
    );
  type tp_numFmts is table of tp_numFmt index by pls_integer;
  type tp_fill is record
    ( patternType varchar2(30)
    , fgRGB varchar2(8)
    );
  type tp_fills is table of tp_fill index by pls_integer;
  type tp_cellXfs is table of tp_xf_fmt index by pls_integer;
  type tp_font is record
    ( name varchar2(100)
    , family pls_integer
    , fontsize number
    , theme pls_integer
    , RGB varchar2(8)
    , underline boolean
    , italic boolean
    , bold boolean
    );
  type tp_fonts is table of tp_font index by pls_integer;
  type tp_border is record
    ( top    varchar2(17)
    , bottom varchar2(17)
    , left   varchar2(17)
    , right  varchar2(17)
    , rgb    varchar2(8)
    );
  type tp_borders is table of tp_border index by pls_integer;
  type tp_numFmtIndexes is table of pls_integer index by pls_integer;
  type tp_strings is table of pls_integer index by varchar2(32767 char);
  type tp_str_ind is table of varchar2(32767 char) index by pls_integer;
  type tp_defined_name is record
    ( name varchar2(32767 char)
    , ref varchar2(32767 char)
    , sheet pls_integer
    );
  type tp_defined_names is table of tp_defined_name index by pls_integer;
  type tp_image is record
    ( img blob
    , hash raw(4)
    , width  pls_integer
    , height pls_integer
    );
  type tp_images is table of tp_image index by pls_integer;
  type tp_book is record
    ( sheets tp_sheets
    , strings tp_strings
    , str_ind tp_str_ind
    , str_cnt pls_integer := 0
    , fonts tp_fonts
    , fills tp_fills
    , borders tp_borders
    , numFmts tp_numFmts
    , cellXfs tp_cellXfs
    , numFmtIndexes tp_numFmtIndexes
    , defined_names tp_defined_names
    , images        tp_images
    , tables        tp_tables
    );
  workbook tp_book;

  --------------------------------------------------------------------------
  -- Package state
  --------------------------------------------------------------------------
  g_useXf boolean := true;

  --------------------------------------------------------------------------
  -- UTF-8 blob buffered-write utilities
  --
  -- Accumulates text fragments in a VARCHAR2 buffer and flushes to a BLOB
  -- only when the buffer overflows. This dramatically reduces the number of
  -- dbms_lob.writeappend calls when building large XML documents.
  --------------------------------------------------------------------------
  g_addtxt2utf8blob_tmp varchar2(32767);
  procedure addtxt2utf8blob_init( p_blob in out nocopy blob )
  is
  begin
    g_addtxt2utf8blob_tmp := null;
    dbms_lob.createtemporary( p_blob, true );
  end;
  procedure addtxt2utf8blob_finish( p_blob in out nocopy blob )
  is
    l_raw raw(32767);
  begin
    l_raw := utl_i18n.string_to_raw( g_addtxt2utf8blob_tmp, 'AL32UTF8' );
    dbms_lob.writeappend( p_blob, utl_raw.length( l_raw ), l_raw );
  exception
    when value_error
    then
      l_raw := utl_i18n.string_to_raw( substr( g_addtxt2utf8blob_tmp, 1, 16381 ), 'AL32UTF8' );
      dbms_lob.writeappend( p_blob, utl_raw.length( l_raw ), l_raw );
      l_raw := utl_i18n.string_to_raw( substr( g_addtxt2utf8blob_tmp, 16382 ), 'AL32UTF8' );
      dbms_lob.writeappend( p_blob, utl_raw.length( l_raw ), l_raw );
  end;
  procedure addtxt2utf8blob( p_txt varchar2, p_blob in out nocopy blob )
  is
  begin
    g_addtxt2utf8blob_tmp := g_addtxt2utf8blob_tmp || p_txt;
  exception
    when value_error
    then
      addtxt2utf8blob_finish( p_blob );
      g_addtxt2utf8blob_tmp := p_txt;
  end;

  --------------------------------------------------------------------------
  -- Low-level helpers
  --------------------------------------------------------------------------

  -- Converts a number to a little-endian RAW of p_bytes length.
  function little_endian( p_big number, p_bytes pls_integer := 4 )
  return raw
  is
  begin
    if p_big < 0
    then
      return utl_raw.reverse( to_char( 4294967296 + p_big, 'fm0XXXXXXX' ) );
    else
      return utl_raw.reverse( to_char( p_big, substr( 'fm0XXXXXXXXXXXXXXXXXXX', 1, 2 + 2 * p_bytes ) ) );
    end if;
  end;

  -- Reads p_len bytes at position p_pos from a BLOB as a little-endian integer.
  function blob2num( p_blob blob, p_len integer, p_pos integer )
  return number
  is
  begin
    return utl_raw.cast_to_binary_integer( dbms_lob.substr( p_blob, p_len, p_pos ), utl_raw.little_endian );
  end;

  --------------------------------------------------------------------------
  -- ZIP archive procedures
  --
  -- Build a ZIP archive incrementally. add1file appends one file entry,
  -- finish_zip writes the central directory and end-of-central-directory
  -- record.  Together they produce a valid ZIP (PKZIP 2.0) container.
  --------------------------------------------------------------------------

  -- Appends a single file to the ZIP archive BLOB.
  -- Compresses with DEFLATE when beneficial, otherwise stores raw.
  procedure add1file
    ( p_zipped_blob in out blob
    , p_name varchar2
    , p_content blob
    )
  is
    l_now date;
    l_blob blob;
    l_len integer;
    l_clen integer;
    l_crc32 raw(4) := hextoraw( '00000000' );
    l_compressed boolean := false;
    l_name raw(32767);
  begin
    l_now := sysdate;
    l_len := nvl( dbms_lob.getlength( p_content ), 0 );
    if l_len > 0
    then
      l_blob := utl_compress.lz_compress( p_content );
      l_clen := dbms_lob.getlength( l_blob ) - 18;
      l_compressed := l_clen < l_len;
      l_crc32 := dbms_lob.substr( l_blob, 4, l_clen + 11 );
    end if;
    if not l_compressed
    then
      l_clen := l_len;
      l_blob := p_content;
    end if;
    if p_zipped_blob is null
    then
      dbms_lob.createtemporary( p_zipped_blob, true );
    end if;
    l_name := utl_i18n.string_to_raw( p_name, 'AL32UTF8' );
    dbms_lob.append( p_zipped_blob
                   , utl_raw.concat( c_LOCAL_FILE_HEADER -- Local file header signature
                                   , hextoraw( '1400' )  -- version 2.0
                                   , case when l_name = utl_i18n.string_to_raw( p_name, 'US8PC437' )
                                       then hextoraw( '0000' ) -- no General purpose bits
                                       else hextoraw( '0008' ) -- set Language encoding flag (EFS)
                                     end
                                   , case when l_compressed
                                        then hextoraw( '0800' ) -- deflate
                                        else hextoraw( '0000' ) -- stored
                                     end
                                   , little_endian( to_number( to_char( l_now, 'ss' ) ) / 2
                                                  + to_number( to_char( l_now, 'mi' ) ) * 32
                                                  + to_number( to_char( l_now, 'hh24' ) ) * 2048
                                                  , 2
                                                  ) -- File last modification time
                                   , little_endian( to_number( to_char( l_now, 'dd' ) )
                                                  + to_number( to_char( l_now, 'mm' ) ) * 32
                                                  + ( to_number( to_char( l_now, 'yyyy' ) ) - 1980 ) * 512
                                                  , 2
                                                  ) -- File last modification date
                                   , l_crc32 -- CRC-32
                                   , little_endian( l_clen )                      -- compressed size
                                   , little_endian( l_len )                       -- uncompressed size
                                   , little_endian( utl_raw.length( l_name ), 2 ) -- File name length
                                   , hextoraw( '0000' )                           -- Extra field length
                                   , l_name                                       -- File name
                                   )
                   );
    if l_compressed
    then
      dbms_lob.copy( p_zipped_blob, l_blob, l_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 11 ); -- compressed content
    elsif l_clen > 0
    then
      dbms_lob.copy( p_zipped_blob, l_blob, l_clen, dbms_lob.getlength( p_zipped_blob ) + 1, 1 ); --  content
    end if;
    if dbms_lob.istemporary( l_blob ) = 1
    then
      dbms_lob.freetemporary( l_blob );
    end if;
  end;

  -- Finalises the ZIP archive by writing the central directory and
  -- end-of-central-directory record.
  procedure finish_zip( p_zipped_blob in out blob )
  is
    l_cnt pls_integer := 0;
    l_offs integer;
    l_offs_dir_header integer;
    l_offs_end_header integer;
    l_comment raw(32767) := utl_raw.cast_to_raw( 'Implementation by Anton Scheffer, ' || c_version );
  begin
    l_offs_dir_header := dbms_lob.getlength( p_zipped_blob );
    l_offs := 1;
    while dbms_lob.substr( p_zipped_blob, utl_raw.length( c_LOCAL_FILE_HEADER ), l_offs ) = c_LOCAL_FILE_HEADER
    loop
      l_cnt := l_cnt + 1;
      dbms_lob.append( p_zipped_blob
                     , utl_raw.concat( hextoraw( '504B0102' )      -- Central directory file header signature
                                     , hextoraw( '1400' )          -- version 2.0
                                     , dbms_lob.substr( p_zipped_blob, 26, l_offs + 4 )
                                     , hextoraw( '0000' )          -- File comment length
                                     , hextoraw( '0000' )          -- Disk number where file starts
                                     , hextoraw( '0000' )          -- Internal file attributes =>
                                                                   --     0000 binary file
                                                                   --     0100 (ascii)text file
                                     , case
                                         when dbms_lob.substr( p_zipped_blob
                                                             , 1
                                                             , l_offs + 30 + blob2num( p_zipped_blob, 2, l_offs + 26 ) - 1
                                                             ) in ( hextoraw( '2F' ) -- /
                                                                  , hextoraw( '5C' ) -- \
                                                                  )
                                         then hextoraw( '10000000' ) -- a directory/folder
                                         else hextoraw( '2000B681' ) -- a file
                                       end                         -- External file attributes
                                     , little_endian( l_offs - 1 ) -- Relative offset of local file header
                                     , dbms_lob.substr( p_zipped_blob
                                                      , blob2num( p_zipped_blob, 2, l_offs + 26 )
                                                      , l_offs + 30
                                                      )            -- File name
                                     )
                     );
      l_offs := l_offs + 30 + blob2num( p_zipped_blob, 4, l_offs + 18 )  -- compressed size
                             + blob2num( p_zipped_blob, 2, l_offs + 26 )  -- File name length
                             + blob2num( p_zipped_blob, 2, l_offs + 28 ); -- Extra field length
    end loop;
    l_offs_end_header := dbms_lob.getlength( p_zipped_blob );
    dbms_lob.append( p_zipped_blob
                   , utl_raw.concat( c_END_OF_CENTRAL_DIRECTORY                                  -- End of central directory signature
                                   , hextoraw( '0000' )                                          -- Number of this disk
                                   , hextoraw( '0000' )                                          -- Disk where central directory starts
                                   , little_endian( l_cnt, 2 )                                   -- Number of central directory records on this disk
                                   , little_endian( l_cnt, 2 )                                   -- Total number of central directory records
                                   , little_endian( l_offs_end_header - l_offs_dir_header )      -- Size of central directory
                                   , little_endian( l_offs_dir_header )                          -- Offset of start of central directory, relative to start of archive
                                   , little_endian( nvl( utl_raw.length( l_comment ), 0 ), 2 )   -- ZIP file comment length
                                   , l_comment
                                   )
                   );
  end;

  --------------------------------------------------------------------------
  -- Column name helper
  --------------------------------------------------------------------------

  -- Converts a 1-based column index to an Excel column letter (A, B, ..., Z, AA, AB, ...).
  function alfan_col( p_col pls_integer )
  return varchar2
  is
  begin
    return case
             when p_col > 702 then chr( 64 + trunc( ( p_col - 27 ) / 676 ) ) || chr( 65 + mod( trunc( ( p_col - 1 ) / 26 ) - 1, 26 ) ) || chr( 65 + mod( p_col - 1, 26 ) )
             when p_col > 26  then chr( 64 + trunc( ( p_col - 1 ) / 26 ) ) || chr( 65 + mod( p_col - 1, 26 ) )
             else chr( 64 + p_col )
           end;
  end;

  --------------------------------------------------------------------------
  -- Workbook management
  --------------------------------------------------------------------------

  -- Resets all workbook state: sheets, strings, fonts, fills, borders,
  -- number formats, cell styles, defined names, tables, and images.
  -- Must be called before building a new workbook.
  procedure clear_workbook
  is
    l_sheet pls_integer;
    l_row_ind pls_integer;
  begin
    l_sheet := workbook.sheets.first;
    while l_sheet is not null
    loop
      l_row_ind := workbook.sheets( l_sheet ).rows.first;
      while l_row_ind is not null
      loop
        workbook.sheets( l_sheet ).rows( l_row_ind ).delete;
        l_row_ind := workbook.sheets( l_sheet ).rows.next( l_row_ind );
      end loop;
      workbook.sheets( l_sheet ).rows.delete;
      workbook.sheets( l_sheet ).widths.delete;
      workbook.sheets( l_sheet ).autofilters.delete;
      workbook.sheets( l_sheet ).hyperlinks.delete;
      workbook.sheets( l_sheet ).col_fmts.delete;
      workbook.sheets( l_sheet ).row_fmts.delete;
      workbook.sheets( l_sheet ).comments.delete;
      workbook.sheets( l_sheet ).mergecells.delete;
      workbook.sheets( l_sheet ).validations.delete;
      workbook.sheets( l_sheet ).drawings.delete;
      l_sheet := workbook.sheets.next( l_sheet );
    end loop;
    workbook.strings.delete;
    workbook.str_ind.delete;
    workbook.fonts.delete;
    workbook.fills.delete;
    workbook.borders.delete;
    workbook.numFmts.delete;
    workbook.cellXfs.delete;
    workbook.defined_names.delete;
    workbook.tables.delete;
    for i in 1 .. workbook.images.count
    loop
      dbms_lob.freetemporary( workbook.images(i).img );
    end loop;
    workbook.images.delete;
    authors.delete;
    workbook := null;
  end;

  -- Sets the worksheet tab color (hex ARGB value).
  procedure set_tabcolor
    ( p_tabcolor varchar2 -- this is a hex ALPHA Red Green Blue value
    , p_sheet pls_integer := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( l_sheet ).tabcolor := substr( p_tabcolor, 1, 8 );
  end;

  -- Adds a new worksheet to the workbook.
  -- Initialises default font (Calibri), fills (none, gray125), and border if
  -- this is the first sheet.
  procedure new_sheet
    ( p_sheetname      varchar2    := null
    , p_tabcolor       varchar2    := null -- this is a hex ALPHA Red Green Blue value
    , p_show_gridlines boolean     := null
    , p_grid_color_idx pls_integer := null -- index in default color palette 0 - 55
    , p_show_headers   boolean     := null
    )
  is
    l_nr pls_integer := workbook.sheets.count + 1;
    l_ind pls_integer;
  begin
    workbook.sheets( l_nr ).name := nvl( dbms_xmlgen.convert( translate( p_sheetname, 'a/\[]*:?', 'a' ) ), 'Sheet' || l_nr );
    workbook.sheets( l_nr ).show_gridlines := p_show_gridlines;
    workbook.sheets( l_nr ).grid_color_idx := p_grid_color_idx;
    workbook.sheets( l_nr ).show_headers   := p_show_headers;
    if workbook.strings.count = 0
    then
     workbook.str_cnt := 0;
    end if;
    if workbook.fonts.count = 0
    then
      l_ind := get_font( 'Calibri' );
    end if;
    if workbook.fills.count = 0
    then
      l_ind := get_fill( 'none' );
      l_ind := get_fill( 'gray125' );
    end if;
    if workbook.borders.count = 0
    then
      l_ind := get_border( '', '', '', '' );
    end if;
    set_tabcolor( p_tabcolor, l_nr );
  end;

  --------------------------------------------------------------------------
  -- Column width helper
  --------------------------------------------------------------------------

  -- Auto-calculates and sets column width based on the number format string.
  -- Uses character count approximation assuming 11pt Calibri.
  procedure set_col_width
    ( p_sheet pls_integer
    , p_col pls_integer
    , p_format varchar2
    )
  is
    l_width number;
    l_nr_chr pls_integer;
  begin
    if p_format is null
    then
      return;
    end if;
    if instr( p_format, ';' ) > 0
    then
      l_nr_chr := length( translate( substr( p_format, 1, instr( p_format, ';' ) - 1 ), 'a\"', 'a' ) );
    else
      l_nr_chr := length( translate( p_format, 'a\"', 'a' ) );
    end if;
    l_width := trunc( ( l_nr_chr * 7 + 5 ) / 7 * 256 ) / 256; -- assume default 11 point Calibri
    if workbook.sheets( p_sheet ).widths.exists( p_col )
    then
      workbook.sheets( p_sheet ).widths( p_col ) :=
        greatest( workbook.sheets( p_sheet ).widths( p_col )
                , l_width
                );
    else
      workbook.sheets( p_sheet ).widths( p_col ) := greatest( l_width, 8.43 );
    end if;
  end;

  --------------------------------------------------------------------------
  -- Format registration
  --
  -- Functions that register and de-duplicate formatting objects (number
  -- formats, fonts, fills, borders, alignments, cell styles) in the
  -- workbook.  Each returns an index that can be passed to cell() or
  -- other APIs.
  --------------------------------------------------------------------------

  -- Converts an Oracle date format mask to its Excel equivalent.
  function OraFmt2Excel( p_format varchar2 := null )
  return varchar2
  is
    l_format varchar2(1000) := substr( p_format, 1, 1000 );
  begin
    l_format := replace( replace( l_format, 'hh24', 'hh' ), 'hh12', 'hh' );
    l_format := replace( l_format, 'mi', 'mm' );
    l_format := replace( replace( replace( l_format, 'AM', '~~' ), 'PM', '~~' ), '~~', 'AM/PM' );
    l_format := replace( replace( replace( l_format, 'am', '~~' ), 'pm', '~~' ), '~~', 'AM/PM' );
    l_format := replace( replace( l_format, 'day', 'DAY' ), 'DAY', 'dddd' );
    l_format := replace( replace( l_format, 'dy', 'DY' ), 'DAY', 'ddd' );
    l_format := replace( replace( l_format, 'RR', 'RR' ), 'RR', 'YY' );
    l_format := replace( replace( l_format, 'month', 'MONTH' ), 'MONTH', 'mmmm' );
    l_format := replace( replace( l_format, 'mon', 'MON' ), 'MON', 'mmm' );
    l_format := replace( l_format, '9', '#' );
    l_format := replace( l_format, 'D', '.' );
    l_format := replace( l_format, 'G', ',' );
    return l_format;
  end;

  -- Registers a custom number format and returns its numFmtId.
  -- Returns 0 when p_format is NULL (= General).
  -- De-duplicates: returns existing ID if format already registered.
  function get_numFmt( p_format varchar2 := null )
  return pls_integer
  is
    l_cnt pls_integer;
    l_numFmtId pls_integer;
  begin
    if p_format is null
    then
      return 0;
    end if;
    l_cnt := workbook.numFmts.count;
    for i in 1 .. l_cnt
    loop
      if workbook.numFmts( i ).formatCode = p_format
      then
        l_numFmtId := workbook.numFmts( i ).numFmtId;
        exit;
      end if;
    end loop;
    if l_numFmtId is null
    then
      l_numFmtId := case when l_cnt = 0 then 164 else workbook.numFmts( l_cnt ).numFmtId + 1 end;
      l_cnt := l_cnt + 1;
      workbook.numFmts( l_cnt ).numFmtId := l_numFmtId;
      workbook.numFmts( l_cnt ).formatCode := p_format;
      workbook.numFmtIndexes( l_numFmtId ) := l_cnt;
    end if;
    return l_numFmtId;
  end;

  -- Registers a font and returns its 0-based index.
  -- De-duplicates: returns existing index if all attributes match.
  function get_font
    ( p_name varchar2
    , p_family pls_integer := 2
    , p_fontsize number := 11
    , p_theme pls_integer := 1
    , p_underline boolean := false
    , p_italic boolean := false
    , p_bold boolean := false
    , p_rgb varchar2 := null -- this is a hex ALPHA Red Green Blue value
    )
  return pls_integer
  is
    l_ind pls_integer;
  begin
    if workbook.fonts.count > 0
    then
      for f in 0 .. workbook.fonts.count - 1
      loop
        if (   workbook.fonts( f ).name = p_name
           and workbook.fonts( f ).family = p_family
           and workbook.fonts( f ).fontsize = p_fontsize
           and workbook.fonts( f ).theme = p_theme
           and workbook.fonts( f ).underline = p_underline
           and workbook.fonts( f ).italic = p_italic
           and workbook.fonts( f ).bold = p_bold
           and ( workbook.fonts( f ).rgb = p_rgb
               or ( workbook.fonts( f ).rgb is null and p_rgb is null )
               )
           )
        then
          return f;
        end if;
      end loop;
    end if;
    l_ind := workbook.fonts.count;
    workbook.fonts( l_ind ).name := p_name;
    workbook.fonts( l_ind ).family := p_family;
    workbook.fonts( l_ind ).fontsize := p_fontsize;
    workbook.fonts( l_ind ).theme := p_theme;
    workbook.fonts( l_ind ).underline := p_underline;
    workbook.fonts( l_ind ).italic := p_italic;
    workbook.fonts( l_ind ).bold := p_bold;
    workbook.fonts( l_ind ).rgb := p_rgb;
    return l_ind;
  end;

  -- Registers a fill pattern and returns its 0-based index.
  function get_fill
    ( p_patternType varchar2
    , p_fgRGB varchar2 := null
    )
  return pls_integer
  is
    l_ind pls_integer;
  begin
    if workbook.fills.count > 0
    then
      for f in 0 .. workbook.fills.count - 1
      loop
        if (   workbook.fills( f ).patternType = p_patternType
           and nvl( workbook.fills( f ).fgRGB, 'x' ) = nvl( upper( p_fgRGB ), 'x' )
           )
        then
          return f;
        end if;
      end loop;
    end if;
    l_ind := workbook.fills.count;
    workbook.fills( l_ind ).patternType := p_patternType;
    workbook.fills( l_ind ).fgRGB := upper( p_fgRGB );
    return l_ind;
  end;

  -- Registers a border style and returns its 0-based index.
  function get_border
    ( p_top    varchar2 := 'thin'
    , p_bottom varchar2 := 'thin'
    , p_left   varchar2 := 'thin'
    , p_right  varchar2 := 'thin'
    , p_rgb    varchar2 := null
    )
  return pls_integer
  is
    l_ind pls_integer;
  begin
    if workbook.borders.count > 0
    then
      for b in 0 .. workbook.borders.count - 1
      loop
        if (   nvl( workbook.borders( b ).top, 'x' )    = nvl( p_top, 'x' )
           and nvl( workbook.borders( b ).bottom, 'x' ) = nvl( p_bottom, 'x' )
           and nvl( workbook.borders( b ).left, 'x' )   = nvl( p_left, 'x' )
           and nvl( workbook.borders( b ).right, 'x' )  = nvl( p_right, 'x' )
           and nvl( workbook.borders( b ).rgb, 'x' )    = nvl( p_rgb, 'x' )
           )
        then
          return b;
        end if;
      end loop;
    end if;
    l_ind := workbook.borders.count;
    workbook.borders( l_ind ).top    := p_top;
    workbook.borders( l_ind ).bottom := p_bottom;
    workbook.borders( l_ind ).left   := p_left;
    workbook.borders( l_ind ).right  := p_right;
    workbook.borders( l_ind ).rgb    := p_rgb;
    return l_ind;
  end;

  -- Builds a tp_alignment record from the given parameters.
  -- Rotation is clamped to 1..180 (NULL if out of range).
  function get_alignment
    ( p_vertical   varchar2 := null
    , p_horizontal varchar2 := null
    , p_wrapText   boolean  := null
    , p_rotation   number   := null
    )
  return tp_alignment
  is
    l_rotation number;
    l_rv tp_alignment;
  begin
    l_rotation := round( p_rotation );
    if l_rotation <= 0 or l_rotation > 180
    then
      l_rotation := null;
    end if;
    l_rv.vertical   := p_vertical;
    l_rv.horizontal := p_horizontal;
    l_rv.wrapText   := p_wrapText;
    l_rv.rotation   := l_rotation;
    return l_rv;
  end;

  -- Registers a combined cell style (XF record) and returns its 1-based index.
  -- Returns NULL when all components are default (no styling needed).
  function get_xfid
    ( p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    )
  return pls_integer
  is
    l_XF   tp_XF_fmt;
    l_cnt  pls_integer;
  begin
    l_XF.numFmtId  := coalesce( p_numFmtId, 0 );
    l_XF.fontId    := coalesce( p_fontId  , 0 );
    l_XF.fillId    := coalesce( p_fillId  , 0 );
    l_XF.borderId  := coalesce( p_borderId, 0 );
    l_XF.alignment := p_alignment;
    if (   l_XF.numFmtId + l_XF.fontId + l_XF.fillId + l_XF.borderId = 0
       and l_XF.alignment.vertical   is null
       and l_XF.alignment.horizontal is null
       and l_XF.alignment.rotation   is null
       and not nvl( l_XF.alignment.wrapText, false )
       )
    then
      return null;
    end if;
    l_cnt := workbook.cellXfs.count;
    for i in 1 .. l_cnt
    loop
      if (   workbook.cellXfs( i ).numFmtId = l_XF.numFmtId
         and workbook.cellXfs( i ).fontId   = l_XF.fontId
         and workbook.cellXfs( i ).fillId   = l_XF.fillId
         and workbook.cellXfs( i ).borderId = l_XF.borderId
         and nvl( workbook.cellXfs( i ).alignment.vertical, 'x' )   = nvl( l_XF.alignment.vertical, 'x' )
         and nvl( workbook.cellXfs( i ).alignment.horizontal, 'x' ) = nvl( l_XF.alignment.horizontal, 'x' )
         and nvl( workbook.cellXfs( i ).alignment.rotation, -7815 ) = nvl( l_XF.alignment.rotation, -7815 )
         and nvl( workbook.cellXfs( i ).alignment.wrapText, false ) = nvl( l_XF.alignment.wrapText, false )
         )
      then
        return i;
      end if;
    end loop;
    l_cnt := l_cnt + 1;
    workbook.cellXfs( l_cnt ) := l_XF;
    return l_cnt;
  end get_XfId;

  -- Resolves the effective cell style by cascading column and row defaults,
  -- then registers the combined XF.  Returns the s="N" attribute string for
  -- the <c> element, or empty string if no style applies.
  function get_XfId
    ( p_sheet     pls_integer
    , p_col       pls_integer
    , p_row       pls_integer
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    )
  return varchar2
  is
    l_XfId   pls_integer;
    l_XF     tp_XF_fmt;
    l_col_XF tp_XF_fmt;
    l_row_XF tp_XF_fmt;
  begin
    if not g_useXf
    then
      return '';
    end if;
    if workbook.sheets( p_sheet ).col_fmts.exists( p_col )
    then
      l_col_XF := workbook.sheets( p_sheet ).col_fmts( p_col );
    end if;
    if workbook.sheets( p_sheet ).row_fmts.exists( p_row )
    then
      l_row_XF := workbook.sheets( p_sheet ).row_fmts( p_row );
    end if;
    l_XF.numFmtId  := coalesce( p_numFmtId, l_col_XF.numFmtId, l_row_XF.numFmtId, 0 );
    l_XF.fontId    := coalesce( p_fontId  , l_col_XF.fontId  , l_row_XF.fontId  , 0 );
    l_XF.fillId    := coalesce( p_fillId  , l_col_XF.fillId  , l_row_XF.fillId  , 0 );
    l_XF.borderId  := coalesce( p_borderId, l_col_XF.borderId, l_row_XF.borderId, 0 );
    l_XF.alignment := get_alignment
                        ( coalesce( p_alignment.vertical, l_col_XF.alignment.vertical, l_row_XF.alignment.vertical )
                        , coalesce( p_alignment.horizontal, l_col_XF.alignment.horizontal, l_row_XF.alignment.horizontal )
                        , coalesce( p_alignment.wrapText, l_col_XF.alignment.wrapText, l_row_XF.alignment.wrapText )
                        , coalesce( p_alignment.rotation, l_col_XF.alignment.rotation, l_row_XF.alignment.rotation )
                        );
    l_xfid := get_xfid( l_XF.numFmtId, l_XF.fontId, l_XF.fillId, l_XF.borderId, l_XF.alignment );
    if l_xfid is null
    then
      return '';
    end if;
    if l_XF.numFmtId > 0
    then
      set_col_width( p_sheet, p_col, workbook.numFmts( workbook.numFmtIndexes( l_XF.numFmtId ) ).formatCode );
    end if;
    return 's="' || l_xfid || '"';
  end get_XfId;

  --------------------------------------------------------------------------
  -- Cell data procedures
  --
  -- Three overloads for NUMBER, VARCHAR2, and DATE values.
  -- Each stores the value in the sheet's row/column matrix and resolves
  -- the cell style via get_XfId (unless an explicit p_xfId is supplied).
  --------------------------------------------------------------------------

  -- Writes a numeric value to a cell.
  procedure cell
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_value number
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).value := p_value;
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).style :=
      case when p_xfid is null
        then get_XfId( l_sheet, p_col, p_row, p_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment )
        else 's="' || p_xfid || '"'
      end;
  end;

  -- Registers a string in the shared strings table and returns its index.
  -- Increments the total string usage counter (str_cnt) on every call.
  function add_string( p_string varchar2 )
  return pls_integer
  is
    l_cnt pls_integer;
  begin
    if workbook.strings.exists( nvl( p_string, '' ) )
    then
      l_cnt := workbook.strings( nvl( p_string, '' ) );
    else
      l_cnt := workbook.strings.count;
      workbook.str_ind( l_cnt ) := p_string;
      workbook.strings( nvl( p_string, '' ) ) := l_cnt;
    end if;
    workbook.str_cnt := workbook.str_cnt + 1;
    return l_cnt;
  end;

  -- Writes a string value to a cell.
  -- Auto-enables wrapText when the string contains CR characters.
  procedure cell
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_value varchar2
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
    l_alignment tp_alignment := p_alignment;
  begin
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).value := add_string( p_value );
    if l_alignment.wrapText is null and instr( p_value, chr(13) ) > 0
    then
      l_alignment.wrapText := true;
    end if;
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).style := 't="s" ' ||
      case when p_xfid is null
        then get_XfId( l_sheet, p_col, p_row, p_numFmtId, p_fontId, p_fillId, p_borderId, l_alignment )
        else 's="' || p_xfid || '"'
      end;
  end;

  -- Writes a date value to a cell.
  -- Converts the date to an Excel serial number (days since 1900-01-01).
  -- Defaults to 'dd/mm/yyyy' format when no numFmtId or column/row format
  -- is specified.
  procedure cell
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_value date
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    )
  is
    l_xfId     varchar2(100);
    l_numFmtId pls_integer := p_numFmtId;
    l_sheet    pls_integer := nvl( p_sheet, workbook.sheets.count );
    l_tmp      number;
  begin
    l_tmp := p_value - date '1900-03-01';
    l_tmp := l_tmp + case when l_tmp < 0 then 60 else 61 end;
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).value := l_tmp;
    if p_xfId is null
    then
      if l_numFmtId is null
         and not (   workbook.sheets( l_sheet ).col_fmts.exists( p_col )
                 and workbook.sheets( l_sheet ).col_fmts( p_col ).numFmtId is not null
                 )
         and not (   workbook.sheets( l_sheet ).row_fmts.exists( p_row )
                 and workbook.sheets( l_sheet ).row_fmts( p_row ).numFmtId is not null
                 )
      then
        l_numFmtId := get_numFmt( 'dd/mm/yyyy' );
      end if;
      l_xfId := get_XfId( l_sheet, p_col, p_row, l_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment );
    else
      l_xfId := 's="' || p_xfid || '"';
    end if;
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).style := l_xfId;
  end;

  --------------------------------------------------------------------------
  -- Formula procedures
  --------------------------------------------------------------------------

  -- Writes a formula that evaluates to a number.
  procedure num_formula
    ( p_col           pls_integer
    , p_row           pls_integer
    , p_formula       varchar2
    , p_default_value number := null
    , p_numFmtId      pls_integer  := null
    , p_fontId        pls_integer  := null
    , p_fillId        pls_integer  := null
    , p_borderId      pls_integer  := null
    , p_alignment     tp_alignment := null
    , p_sheet         pls_integer  := null
    , p_xfId          pls_integer  := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    cell( p_col, p_row, p_default_value, p_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment, l_sheet, p_xfId );
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).formula := '<f>' || p_formula || '</f>';
  end num_formula;

  -- Writes a formula that evaluates to a string.
  procedure str_formula
    ( p_col           pls_integer
    , p_row           pls_integer
    , p_formula       varchar2
    , p_default_value varchar2 := null
    , p_numFmtId      pls_integer  := null
    , p_fontId        pls_integer  := null
    , p_fillId        pls_integer  := null
    , p_borderId      pls_integer  := null
    , p_alignment     tp_alignment := null
    , p_sheet         pls_integer  := null
    , p_xfId          pls_integer  := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    cell( p_col, p_row, p_default_value, p_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment, l_sheet, p_xfId );
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).formula := '<f>' || p_formula || '</f>';
  end str_formula;

  -- Writes a formula that evaluates to a date.
  procedure date_formula
    ( p_col           pls_integer
    , p_row           pls_integer
    , p_formula       varchar2
    , p_default_value date := null
    , p_numFmtId      pls_integer  := null
    , p_fontId        pls_integer  := null
    , p_fillId        pls_integer  := null
    , p_borderId      pls_integer  := null
    , p_alignment     tp_alignment := null
    , p_sheet         pls_integer  := null
    , p_xfId          pls_integer  := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    cell( p_col, p_row, p_default_value, p_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment, l_sheet, p_xfId );
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).formula := '<f>' || p_formula || '</f>';
  end date_formula;

  -- Internal: writes a date cell with a pre-resolved XfId string.
  -- Used by query2sheet when g_useXf is false.
  procedure query_date_cell
    ( p_col pls_integer
    , p_row pls_integer
    , p_value date
    , p_sheet pls_integer := null
    , p_XfId varchar2
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    cell( p_col, p_row, p_value, 0, p_sheet => l_sheet );
    workbook.sheets( l_sheet ).rows( p_row )( p_col ).style := p_XfId;
  end;

  --------------------------------------------------------------------------
  -- Cell features (hyperlinks, comments, merged cells, validations)
  --------------------------------------------------------------------------

  -- Adds a clickable hyperlink to a cell.
  -- Either p_url (external) or p_location (internal bookmark) must be set.
  procedure hyperlink
    ( p_col pls_integer
    , p_row pls_integer
    , p_url varchar2 := null
    , p_value varchar2 := null
    , p_sheet pls_integer := null
    , p_location varchar2 := null
    , p_tooltip varchar2 := null
    )
  is
    l_ind pls_integer;
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    if p_url is not null or p_location is not null
    then
      workbook.sheets( l_sheet ).rows( p_row )( p_col ).value := add_string( coalesce( p_value, p_url, p_location ) );
      workbook.sheets( l_sheet ).rows( p_row )( p_col ).style := 't="s" ' || get_XfId( l_sheet, p_col, p_row, '', get_font( 'Calibri', p_theme => 10, p_underline => true ) );
      l_ind := workbook.sheets( l_sheet ).hyperlinks.count + 1;
      workbook.sheets( l_sheet ).hyperlinks( l_ind ).cell := alfan_col( p_col ) || p_row;
      workbook.sheets( l_sheet ).hyperlinks( l_ind ).url := p_url;
      workbook.sheets( l_sheet ).hyperlinks( l_ind ).location := p_location;
      workbook.sheets( l_sheet ).hyperlinks( l_ind ).tooltip := p_tooltip;
    end if;
  end;

  -- Adds a comment (note) to a cell.
  procedure comment
    ( p_col pls_integer
    , p_row pls_integer
    , p_text varchar2
    , p_author varchar2 := null
    , p_width pls_integer := 150
    , p_height pls_integer := 100
    , p_sheet pls_integer := null
    )
  is
    l_ind pls_integer;
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    l_ind := workbook.sheets( l_sheet ).comments.count + 1;
    workbook.sheets( l_sheet ).comments( l_ind ).row := p_row;
    workbook.sheets( l_sheet ).comments( l_ind ).column := p_col;
    workbook.sheets( l_sheet ).comments( l_ind ).text := dbms_xmlgen.convert( p_text );
    workbook.sheets( l_sheet ).comments( l_ind ).author := dbms_xmlgen.convert( p_author );
    workbook.sheets( l_sheet ).comments( l_ind ).width := p_width;
    workbook.sheets( l_sheet ).comments( l_ind ).height := p_height;
  end;

  -- Merges a rectangular range of cells.
  procedure mergecells
    ( p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_sheet pls_integer := null
    )
  is
    l_ind pls_integer;
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    l_ind := workbook.sheets( l_sheet ).mergecells.count + 1;
    workbook.sheets( l_sheet ).mergecells( l_ind ) := alfan_col( p_tl_col ) || p_tl_row || ':' || alfan_col( p_br_col ) || p_br_row;
  end;

  -- Internal: adds a data validation rule to a sheet.
  procedure add_validation
    ( p_type varchar2
    , p_sqref varchar2
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_formula1 varchar2 := null
    , p_formula2 varchar2 := null
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    )
  is
    l_ind pls_integer;
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    l_ind := workbook.sheets( l_sheet ).validations.count + 1;
    workbook.sheets( l_sheet ).validations( l_ind ).type := p_type;
    workbook.sheets( l_sheet ).validations( l_ind ).errorstyle := p_style;
    workbook.sheets( l_sheet ).validations( l_ind ).sqref := p_sqref;
    workbook.sheets( l_sheet ).validations( l_ind ).formula1 := p_formula1;
    workbook.sheets( l_sheet ).validations( l_ind ).error_title := p_error_title;
    workbook.sheets( l_sheet ).validations( l_ind ).error_txt := p_error_txt;
    workbook.sheets( l_sheet ).validations( l_ind ).title := p_title;
    workbook.sheets( l_sheet ).validations( l_ind ).prompt := p_prompt;
    workbook.sheets( l_sheet ).validations( l_ind ).showerrormessage := p_show_error;
  end;

  -- Adds a dropdown list validation referencing a cell range.
  procedure list_validation
    ( p_sqref_col pls_integer
    , p_sqref_row pls_integer
    , p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    )
  is
  begin
    add_validation( 'list'
                  , alfan_col( p_sqref_col ) || p_sqref_row
                  , p_style => lower( p_style )
                  , p_formula1 => '$' || alfan_col( p_tl_col ) || '$' ||  p_tl_row || ':$' || alfan_col( p_br_col ) || '$' || p_br_row
                  , p_title => p_title
                  , p_prompt => p_prompt
                  , p_show_error => p_show_error
                  , p_error_title => p_error_title
                  , p_error_txt => p_error_txt
                  , p_sheet => p_sheet
                  );
  end;

  -- Adds a dropdown list validation referencing a defined name.
  procedure list_validation
    ( p_sqref_col pls_integer
    , p_sqref_row pls_integer
    , p_defined_name varchar2
    , p_style varchar2 := 'stop' -- stop, warning, information
    , p_title varchar2 := null
    , p_prompt varchar := null
    , p_show_error boolean := false
    , p_error_title varchar2 := null
    , p_error_txt varchar2 := null
    , p_sheet pls_integer := null
    )
  is
  begin
    add_validation( 'list'
                  , alfan_col( p_sqref_col ) || p_sqref_row
                  , p_style => lower( p_style )
                  , p_formula1 => p_defined_name
                  , p_title => p_title
                  , p_prompt => p_prompt
                  , p_show_error => p_show_error
                  , p_error_title => p_error_title
                  , p_error_txt => p_error_txt
                  , p_sheet => p_sheet
                  );
  end;

  -- Creates a named range pointing to a cell range on a sheet.
  procedure defined_name
    ( p_tl_col pls_integer -- top left
    , p_tl_row pls_integer
    , p_br_col pls_integer -- bottom right
    , p_br_row pls_integer
    , p_name varchar2
    , p_sheet pls_integer := null
    , p_localsheet pls_integer := null
    )
  is
    l_ind pls_integer;
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    l_ind := workbook.defined_names.count + 1;
    workbook.defined_names( l_ind ).name := p_name;
    workbook.defined_names( l_ind ).ref := 'Sheet' || l_sheet || '!$' || alfan_col( p_tl_col ) || '$' ||  p_tl_row || ':$' || alfan_col( p_br_col ) || '$' || p_br_row;
    workbook.defined_names( l_ind ).sheet := p_localsheet;
  end;

  --------------------------------------------------------------------------
  -- Column & row formatting
  --------------------------------------------------------------------------

  -- Sets an explicit column width (in character units).
  procedure set_column_width
    ( p_col pls_integer
    , p_width number
    , p_sheet pls_integer := null
    )
  is
    l_width number;
  begin
    l_width := trunc( round( p_width * 7 ) * 256 / 7 ) / 256;
    workbook.sheets( nvl( p_sheet, workbook.sheets.count ) ).widths( p_col ) := l_width;
  end;

  -- Sets default formatting for an entire column.
  procedure set_column
    ( p_col pls_integer
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( l_sheet ).col_fmts( p_col ).numFmtId := p_numFmtId;
    workbook.sheets( l_sheet ).col_fmts( p_col ).fontId := p_fontId;
    workbook.sheets( l_sheet ).col_fmts( p_col ).fillId := p_fillId;
    workbook.sheets( l_sheet ).col_fmts( p_col ).borderId := p_borderId;
    workbook.sheets( l_sheet ).col_fmts( p_col ).alignment := p_alignment;
  end;

  -- Sets default formatting and/or height for an entire row.
  procedure set_row
    ( p_row pls_integer
    , p_numFmtId pls_integer := null
    , p_fontId pls_integer := null
    , p_fillId pls_integer := null
    , p_borderId pls_integer := null
    , p_alignment tp_alignment := null
    , p_sheet pls_integer := null
    , p_height number := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
    l_cells tp_cells;
  begin
    workbook.sheets( l_sheet ).row_fmts( p_row ).numFmtId := p_numFmtId;
    workbook.sheets( l_sheet ).row_fmts( p_row ).fontId := p_fontId;
    workbook.sheets( l_sheet ).row_fmts( p_row ).fillId := p_fillId;
    workbook.sheets( l_sheet ).row_fmts( p_row ).borderId := p_borderId;
    workbook.sheets( l_sheet ).row_fmts( p_row ).alignment := p_alignment;
    workbook.sheets( l_sheet ).row_fmts( p_row ).height := trunc( p_height * 4 / 3 ) * 3 / 4;
    if not workbook.sheets( l_sheet ).rows.exists( p_row )
    then
      workbook.sheets( l_sheet ).rows( p_row ) := l_cells;
    end if;
  end;

  --------------------------------------------------------------------------
  -- Freeze panes
  --------------------------------------------------------------------------

  -- Freezes the top N rows.
  procedure freeze_rows
    ( p_nr_rows pls_integer := 1
    , p_sheet pls_integer := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( l_sheet ).freeze_cols := null;
    workbook.sheets( l_sheet ).freeze_rows := p_nr_rows;
  end;

  -- Freezes the left N columns.
  procedure freeze_cols
    ( p_nr_cols pls_integer := 1
    , p_sheet pls_integer := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( l_sheet ).freeze_rows := null;
    workbook.sheets( l_sheet ).freeze_cols := p_nr_cols;
  end;

  -- Freezes both rows and columns at the given position.
  procedure freeze_pane
    ( p_col pls_integer
    , p_row pls_integer
    , p_sheet pls_integer := null
    )
  is
    l_sheet pls_integer := nvl( p_sheet, workbook.sheets.count );
  begin
    workbook.sheets( l_sheet ).freeze_rows := p_row;
    workbook.sheets( l_sheet ).freeze_cols := p_col;
  end;

  --------------------------------------------------------------------------
  -- Autofilter & tables
  --------------------------------------------------------------------------

  -- Enables autofilter dropdown buttons on a range.
  procedure set_autofilter
    ( p_column_start pls_integer := null
    , p_column_end pls_integer := null
    , p_row_start pls_integer := null
    , p_row_end pls_integer := null
    , p_sheet pls_integer := null
    )
  is
    l_ind pls_integer;
    l_sheet pls_integer := coalesce( p_sheet, workbook.sheets.count );
  begin
    l_ind := 1;
    workbook.sheets( l_sheet ).autofilters( l_ind ).column_start := p_column_start;
    workbook.sheets( l_sheet ).autofilters( l_ind ).column_end := p_column_end;
    workbook.sheets( l_sheet ).autofilters( l_ind ).row_start := p_row_start;
    workbook.sheets( l_sheet ).autofilters( l_ind ).row_end := p_row_end;
    defined_name
      ( p_column_start
      , p_row_start
      , p_column_end
      , p_row_end
      , '_xlnm._FilterDatabase'
      , l_sheet
      , l_sheet - 1
      );
  end;

  -- Defines a formatted table on a range with a named style.
  procedure set_table
    ( p_column_start pls_integer
    , p_column_end   pls_integer
    , p_row_start    pls_integer
    , p_row_end      pls_integer
    , p_style        varchar2
    , p_name         varchar2    := null
    , p_sheet        pls_integer := null
    )
  is
    l_table tp_table;
    l_cnt pls_integer := workbook.tables.count + 1;
  begin
    l_table.sheet := coalesce( p_sheet, workbook.sheets.count );
    l_table.column_start := p_column_start;
    l_table.column_end   := p_column_end;
    l_table.row_start    := p_row_start;
    l_table.row_end      := p_row_end;
    l_table.name         := coalesce( p_name, 'Table' || l_cnt );
    l_table.style        := p_style;
    workbook.tables( l_cnt ) := l_table;
  end;

  --------------------------------------------------------------------------
  -- XML helper procedures
  --
  -- add1xml converts a CLOB of XML to UTF-8 BLOB and adds it as a file
  -- to the ZIP archive.
  --------------------------------------------------------------------------

  -- Converts an XML CLOB to UTF-8 and adds it as a ZIP entry.
  procedure add1xml
    ( p_excel in out nocopy blob
    , p_filename varchar2
    , p_xml clob
    )
  is
    l_tmp blob;
    l_dest_offset integer := 1;
    l_src_offset integer := 1;
    l_lang_context integer;
    l_warning integer;
  begin
    l_lang_context := dbms_lob.DEFAULT_LANG_CTX;
    dbms_lob.createtemporary( l_tmp, true );
    dbms_lob.converttoblob
      ( l_tmp
      , p_xml
      , dbms_lob.lobmaxsize
      , l_dest_offset
      , l_src_offset
      ,  nls_charset_id( 'AL32UTF8'  )
      , l_lang_context
      , l_warning
      );
    add1file( p_excel, p_filename, l_tmp );
    dbms_lob.freetemporary( l_tmp );
  end;

  --------------------------------------------------------------------------
  -- Drawing helper
  --------------------------------------------------------------------------

  -- Generates the DrawingML XML for a single two-cell-anchored image.
  -- Calculates the end column/row and offsets based on image dimensions,
  -- column widths, and row heights.
  function finish_drawing( p_drawing tp_drawing, p_idx pls_integer, p_sheet pls_integer )
  return varchar2
  is
    l_rv varchar2(32767);
    l_col pls_integer;
    l_row pls_integer;
    l_width number;
    l_height number;
    l_col_offs number;
    l_row_offs number;
    l_col_width number;
    l_row_height number;
    l_widths tp_widths;
    l_heights tp_row_fmts;
  begin
    l_width  := workbook.images( p_drawing.img_id ).width;
    l_height := workbook.images( p_drawing.img_id ).height;
    if p_drawing.scale is not null
    then
      l_width  := p_drawing.scale * l_width;
      l_height := p_drawing.scale * l_height;
    end if;
    if workbook.sheets( p_sheet ).widths.count = 0
    then
-- assume default column widths!
-- 64 px = 1 col = 609600
      l_col := trunc( l_width / 64 );
      l_col_offs := ( l_width - l_col * 64 ) * 9525;
      l_col := p_drawing.col - 1 + l_col;
    else
      l_widths := workbook.sheets( p_sheet ).widths;
      l_col := p_drawing.col;
      loop
        if l_widths.exists( l_col )
        then
          l_col_width := round( 7 * l_widths( l_col ) );
        else
          l_col_width := 64;
        end if;
        exit when l_width < l_col_width;
        l_col := l_col + 1;
        l_width := l_width - l_col_width;
      end loop;
      l_col := l_col - 1;
      l_col_offs := l_width * 9525;
    end if;
--
    if workbook.sheets( p_sheet ).row_fmts.count = 0
    then
-- assume default row heigths!
-- 20 px = 1 row = 190500
      l_row := trunc( l_height / 20 );
      l_row_offs := ( l_height - l_row * 20 ) * 9525;
      l_row := p_drawing.row - 1 + l_row;
    else
      l_heights := workbook.sheets( p_sheet ).row_fmts;
      l_row := p_drawing.row;
      loop
        if l_heights.exists( l_row ) and l_heights( l_row ).height is not null
        then
          l_row_height := l_heights( l_row ).height;
          l_row_height := round( 4 * l_row_height / 3 );
        else
          l_row_height := 20;
        end if;
        exit when l_height < l_row_height;
        l_row := l_row + 1;
        l_height := l_height - l_row_height;
      end loop;
      l_row_offs := l_height * 9525;
      l_row := l_row - 1;
    end if;
    l_rv := '<xdr:twoCellAnchor editAs="oneCell">
<xdr:from>
<xdr:col>' || ( p_drawing.col - 1 ) || '</xdr:col>
<xdr:colOff>0</xdr:colOff>
<xdr:row>' || ( p_drawing.row - 1 ) || '</xdr:row>
<xdr:rowOff>0</xdr:rowOff>
</xdr:from>
<xdr:to>
<xdr:col>' || l_col || '</xdr:col>
<xdr:colOff>' || l_col_offs || '</xdr:colOff>
<xdr:row>' || l_row || '</xdr:row>
<xdr:rowOff>' || l_row_offs || '</xdr:rowOff>
</xdr:to>
<xdr:pic>
<xdr:nvPicPr>
<xdr:cNvPr id="3" name="' || coalesce( p_drawing.name, 'Picture ' || p_idx ) || '"';
    if p_drawing.title is not null
    then
      l_rv := l_rv || ' title="' || p_drawing.title || '"';
    end if;
    if p_drawing.description is not null
    then
      l_rv := l_rv || ' descr="' || p_drawing.description || '"';
    end if;
    l_rv := l_rv || '/>
<xdr:cNvPicPr>
<a:picLocks noChangeAspect="1"/>
</xdr:cNvPicPr>
</xdr:nvPicPr>
<xdr:blipFill>
<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' || p_drawing.img_id || '">
<a:extLst>
<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
</a:ext>
</a:extLst>
</a:blip>
<a:stretch>
<a:fillRect/>
</a:stretch>
</xdr:blipFill>
<xdr:spPr>
<a:prstGeom prst="rect">
</a:prstGeom>
</xdr:spPr>
</xdr:pic>
<xdr:clientData/>
</xdr:twoCellAnchor>
';
    return l_rv;
  end;

  -- Returns an XML element with an rgb attribute, or NULL when p_rgb is NULL.
  function add_rgb( p_rgb varchar2, p_tag varchar2 := 'color' )
  return varchar2
  is
  begin
    return case when p_rgb is not null then '<' || p_tag || ' rgb="' || p_rgb || '"/>' end;
  end;

  --------------------------------------------------------------------------
  -- Encryption (conditional compilation)
  --
  -- Only compiled when pck_as_xlsx.use_dbms_crypto = true in the spec.
  -- Implements ECMA-376 Agile Encryption using AES-256-CBC and SHA-1,
  -- wrapped in a Compound File Binary (CFB) container.
  --------------------------------------------------------------------------
$IF pck_as_xlsx.use_dbms_crypto
$THEN
  function excel_encrypt( p_xlsx blob, p_password varchar2 )
  return blob
  is
    t_EncryptionInfo raw(32767);
    t_EncryptedPackage blob;
    --
    c_Free_SecID         constant pls_integer := -1; -- Free sector, may exist in the file, but is not part of any stream
    c_End_Of_Chain_SecID constant pls_integer := -2; -- Trailing SecID in a SecID chain
    c_SAT_SecID          constant pls_integer := -3; -- Sector is used by the sector allocation table
    c_MSAT_SecID         constant pls_integer := -4; -- Sector is used by the master sector allocation table
    --
    c_CLR_Red      constant raw(1) := hextoraw( '00' ); -- Red
    c_CLR_Black    constant raw(1) := hextoraw( '01' ); -- Black
    --
    c_DIR_Empty    constant raw(1) := hextoraw( '00' ); -- Empty
    c_DIR_Storage  constant raw(1) := hextoraw( '01' ); -- User storage
    c_DIR_Stream   constant raw(1) := hextoraw( '02' ); -- User stream
    c_DIR_Lock     constant raw(1) := hextoraw( '03' ); -- LockBytes
    c_DIR_Property constant raw(1) := hextoraw( '04' ); -- Property
    c_DIR_Root     constant raw(1) := hextoraw( '05' ); -- Root storage
    --
    c_Primary raw(200) := hextoraw( '58000000010000004C0000007B00460046003900410033004600300033002D0035003600450046002D0034003600310033002D0042004400440035002D003500410034003100430031004400300037003200340036007D004E0000004D006900630072006F0073006F00660074002E0043006F006E007400610069006E00650072002E0045006E006300720079007000740069006F006E005400720061006E00730066006F0072006D00000001000000010000000100000000000000000000000000000004000000' );
    c_StrongEncryptionDataSpace raw(64) := hextoraw( '0800000001000000320000005300740072006F006E00670045006E006300720079007000740069006F006E005400720061006E00730066006F0072006D000000' );
    c_DataSpaceMap raw(112) := hextoraw( '08000000010000006800000001000000000000002000000045006E0063007200790070007400650064005000610063006B00610067006500320000005300740072006F006E00670045006E006300720079007000740069006F006E004400610074006100530070006100630065000000' );
    c_Version raw(76) := hextoraw( '3C0000004D006900630072006F0073006F00660074002E0043006F006E007400610069006E00650072002E004400610074006100530070006100630065007300010000000100000001000000' );
    --
    type tp_childs is table of pls_integer index by pls_integer;
    type tp_dir_entry is record
      ( rname raw(64)
      , tp_entry raw(1)
      , colour   raw(1) := c_CLR_Red
      , left  pls_integer := -1
      , right pls_integer := -1
      , root  pls_integer := -1
      , childs tp_childs
      , len pls_integer := 0
      , first_sector pls_integer := 0
      );
    type tp_dir is table of tp_dir_entry index by pls_integer;
    t_dir tp_dir;
    t_cf blob;
    t_short_stream blob;
    t_root     pls_integer;
    t_dummy    pls_integer;
    t_storage  pls_integer;
    t_storage2 pls_integer;
    t_sorted boolean;
    t_tmp pls_integer;
    t_header raw(512);
    t_ssz        pls_integer := 512; -- sector size
    t_sssz       pls_integer := 64;  -- short sector size
    t_ss_cutoff  pls_integer := 4096;
    t_sectId     pls_integer;
    t_tmp_sectId t_sectId%type;
    type tp_secids is table of t_sectId%type index by pls_integer;
    t_msat tp_secids;
    t_sat  tp_secids;
    t_ssat tp_secids;
    t_st_dir  pls_integer;
    t_st_ssf  pls_integer;
    t_cnt_ssf pls_integer;
    --
    function is_less( p1 tp_dir_entry, p2 tp_dir_entry )
    return boolean
    is
    begin
      return case sign( utl_raw.length( p1.rname ) - utl_raw.length( p2.rname ) )
               when -1 then true
               when  1 then false
               else upper( utl_i18n.raw_to_char( p1.rname, 'AL16UTF16LE' ) )
                  < upper( utl_i18n.raw_to_char( p2.rname, 'AL16UTF16LE' ) )
             end;
    end;
  --
  function add_dir_entry( p_name varchar2, tp_entry raw, p_parent pls_integer := null, p_stream blob := null, p_prefix raw := null )
  return pls_integer
  is
    t_id pls_integer;
    t_entry tp_dir_entry;
  begin
    t_id := t_dir.count;
    t_entry.tp_entry := tp_entry;
    t_entry.rname := utl_raw.concat( p_prefix, utl_i18n.string_to_raw( p_name, 'AL16UTF16LE' ) );
    if p_parent is not null
    then
      t_dir( p_parent ).childs( t_dir( p_parent ).childs.count ) := t_id;
    end if;
    if tp_entry = c_DIR_Stream
    then
      t_entry.len := dbms_lob.getlength( p_stream );
      if t_entry.len >= t_ss_cutoff
      then
        dbms_lob.append( t_cf, p_stream );
        if mod( t_entry.len, t_ssz ) > 0
        then
          dbms_lob.writeappend( t_cf, t_ssz - mod( t_entry.len, t_ssz ), utl_raw.copies( '00', t_ssz ) );
        end if;
        t_entry.first_sector := t_sat.count;
        for i in t_sat.count .. t_sat.count + trunc( ( t_entry.len - 1 ) / t_ssz ) - 1
        loop
          t_sat( i ) := i + 1;
        end loop;
        t_sat( t_sat.count ) := c_End_Of_Chain_SecID;
      else
        dbms_lob.append( t_short_stream, p_stream );
        if mod( t_entry.len, t_sssz ) > 0
        then
          dbms_lob.writeappend( t_short_stream, t_sssz - mod( t_entry.len, t_sssz ), utl_raw.copies( '00', t_sssz ) );
        end if;
        t_entry.first_sector := t_ssat.count;
        for i in t_ssat.count .. t_ssat.count + trunc( ( t_entry.len - 1 ) / t_sssz ) - 1
        loop
          t_ssat( i ) := i + 1;
        end loop;
        t_ssat( t_ssat.count ) := c_End_Of_Chain_SecID;
      end if;
    end if;
    t_dir( t_id ) := t_entry;
    return t_id;
  end;
--
  procedure doEncryption
    ( p_pw      varchar2
    , p_excel   blob
    , p_package in out blob
    , p_info    in out raw
    )
  is
    c_algo      constant pls_integer   := dbms_crypto.ENCRYPT_AES + dbms_crypto.CHAIN_CBC + dbms_crypto.PAD_ZERO;
    c_keybits   constant pls_integer   := 256 / 8;
    c_hash      constant pls_integer   := dbms_crypto.hash_sh1;
    c_hmac      constant pls_integer   := dbms_crypto.hmac_sh1;
    c_hash_algo constant varchar2(10)  := 'SHA1';
    c_hash_len  constant pls_integer   := utl_raw.length( dbms_crypto.hash( '00', c_hash ) );
    blockSize pls_integer := 16;
    c_spinCount constant pls_integer := 1000;
    c_saltSize  constant pls_integer   := 16;
    c_salt      constant raw(3999)     := dbms_crypto.randombytes( c_saltsize );
    c_data_salt constant raw(3999)     := dbms_crypto.randombytes( c_saltsize );
    c_pw        constant raw(32767)    := utl_i18n.string_to_raw( p_pw, 'AL16UTF16LE' );
    --
    encrVerifierHashInputBlockKey  constant raw(8) :=  hextoraw( 'fea7d2763b4b9e79' );
    encrVerifierHashValueBlockKey  constant raw(8) :=  hextoraw( 'd7aa0f6d3061344e' );
    encryptedKeyValueBlockKey      constant raw(8) :=  hextoraw( '146e0be7abacd0d6' );
    encrIntegritySaltBlockKey      constant raw(8) :=  hextoraw( '5fb2ad010cb9e1f6' );
    encrIntegrityHmacValueBlocKkey constant raw(8) :=  hextoraw( 'a0677f02b22c8433' );
    --
    l_len        integer;
    l_last_block pls_integer;
  decryptedVerifierInput raw(100);
  decryptedVerifierValue raw(100);
  verifierInputKey raw(100);
  verifierValueKey raw(100);
    encryptedKeyValue      varchar2(100);
    encryptedHmacKey       varchar2(100);
    encryptedHmacValue     varchar2(100);
    encryptedVerifierInput varchar2(100);
    encryptedVerifierValue varchar2(100);
    saltRaw           raw(100);
    hashRaw           raw(100);
    ivRaw             raw(100);
    mac               raw(100);
    rkey              raw(100);
    rinp              raw(100);
    decryptedKeyValue raw(100);
    t_block raw(4096);
--
    function GenerateKey( salt raw, password raw, blockKey raw, hashSize number )
    return raw
    is
      hashBuf raw(1000);
    begin
      hashBuf := dbms_crypto.hash( utl_raw.concat( salt, password ), c_hash );
      for i in 0 .. c_spinCount - 1
      loop
        hashBuf := dbms_crypto.hash( utl_raw.concat( little_endian( i ), hashBuf ), c_hash );
      end loop;
      hashBuf := dbms_crypto.hash( utl_raw.concat( hashBuf, blockKey ), c_hash );
      if c_hash_len < hashSize
      then
        hashBuf := utl_raw.concat( hashBuf, utl_raw.copies( hextoraw( '36' ), hashSize ) );
      end if;
      return utl_raw.substr( hashBuf, 1, hashSize );
    end GenerateKey;
  begin
    decryptedKeyValue := dbms_crypto.randombytes( c_keybits );
    rkey := GenerateKey( c_salt, c_pw, encryptedKeyValueBlockKey, c_keybits );
    ivRaw := dbms_crypto.encrypt( decryptedKeyValue, c_algo, rkey, c_salt );
    encryptedKeyValue := utl_raw.cast_to_varchar2( utl_encode.base64_encode( ivRaw ) );
    l_len := dbms_lob.getlength( p_excel );
    p_package := little_endian( l_len, 8 );
    l_last_block := trunc( ( l_len - 1 ) / 4096 );
    for i in 0 .. l_last_block
    loop
      ivRaw := dbms_crypto.hash( utl_raw.concat( c_data_salt, little_endian( i ) ), c_hash );
      if c_hash_len < blockSize
      then
        ivRaw := utl_raw.concat( ivRaw, utl_raw.copies( hextoraw( '36' ), blockSize ) );
      end if;
      ivRaw := utl_raw.substr( ivRaw, 1, blockSize );
      t_block := dbms_lob.substr( p_excel, 4096, 1 + i * 4096 );
      if i = l_last_block and mod( utl_raw.length( t_block ), blockSize ) != 0
      then
        t_block := utl_raw.concat( t_block, utl_raw.copies( 'FF', blockSize - mod( utl_raw.length( t_block ), blockSize ) ) );
      end if;
      dbms_lob.append( p_package, dbms_crypto.encrypt( t_block, c_algo, decryptedKeyValue, ivRaw ) );
    end loop;
--
    saltRaw := dbms_crypto.randombytes( c_hash_len );
    mac := dbms_crypto.mac( p_package, c_hmac, saltRaw );
    ivRaw := dbms_crypto.hash( utl_raw.concat( c_data_salt, encrIntegritySaltBlockKey ), c_hash );
    if utl_raw.length( ivRaw ) < blockSize
    then
      ivRaw := utl_raw.concat( ivRaw, utl_raw.copies( hextoraw( '00' ), blockSize ) );
    end if;
    ivRaw := utl_raw.substr( ivRaw, 1, blockSize );
    saltRaw := dbms_crypto.encrypt( saltRaw, c_algo, decryptedKeyValue, ivRaw );
    encryptedHmacKey := utl_raw.cast_to_varchar2( utl_encode.base64_encode( saltRaw ) );
    ivRaw := dbms_crypto.hash( utl_raw.concat( c_data_salt, encrIntegrityHmacValueBlocKkey ), c_hash );
    if utl_raw.length( ivRaw ) < blockSize
    then
      ivRaw := utl_raw.concat( ivRaw, utl_raw.copies( hextoraw( '00' ), blockSize ) );
    end if;
    ivRaw := utl_raw.substr( ivRaw, 1, blockSize );
    hashRaw := dbms_crypto.encrypt( mac, c_algo, decryptedKeyValue, ivRaw );
    encryptedHmacValue := utl_raw.cast_to_varchar2( utl_encode.base64_encode( hashRaw ) );
--
    rinp := dbms_crypto.randombytes( c_saltSize );
    rkey := GenerateKey( c_salt, c_pw, encrVerifierHashInputBlockKey, c_keybits );
    hashRaw := dbms_crypto.encrypt( rinp, c_algo, rkey, c_salt );
    encryptedVerifierInput := utl_raw.cast_to_varchar2( utl_encode.base64_encode( hashRaw ) );
    rkey := GenerateKey( c_salt, c_pw, encrVerifierHashValueBlockKey, c_keybits );
    rinp := dbms_crypto.hash( rinp, c_hash );
    hashRaw := dbms_crypto.encrypt( rinp, c_algo, rkey, c_salt );
    encryptedVerifierValue := utl_raw.cast_to_varchar2( utl_encode.base64_encode( hashRaw ) );
--
    p_info := utl_raw.concat( hextoraw( '0400040040000000' )
                            , utl_raw.cast_to_raw(
'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' || chr(13) || chr(10) ||
'<encryption' ||
' xmlns="http://schemas.microsoft.com/office/2006/encryption"' ||
' xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><keyData' ||
' saltSize="' || to_char( c_saltsize ) || '"' ||
' blockSize="' || to_char( blocksize ) || '"' ||
' keyBits="' || to_char( c_keybits * 8 )|| '"' ||
' hashSize="' || to_char( c_hash_len ) || '"' ||
' cipherAlgorithm="AES"' ||
' cipherChaining="ChainingModeCBC"' ||
' hashAlgorithm="' || c_hash_algo || '"' ||
' saltValue="' || utl_raw.cast_to_varchar2( utl_encode.base64_encode( c_data_salt ) ) || '"/><dataIntegrity' ||
' encryptedHmacKey="' || encryptedHmacKey || '"' ||
' encryptedHmacValue="' || encryptedHmacValue || '"/><keyEncryptors><keyEncryptor' ||
' uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey' ||
' spinCount="' || to_char( c_spincount ) || '"' ||
' saltSize="' || to_char( c_saltsize ) || '"' ||
' blockSize="' || to_char( blocksize ) || '"' ||
' keyBits="' || to_char( c_keybits * 8 ) || '"' ||
' hashSize="' || to_char( c_hash_len ) || '"' ||
' cipherAlgorithm="AES"' ||
' cipherChaining="ChainingModeCBC"' ||
' hashAlgorithm="' || c_hash_algo || '"' ||
' saltValue="' || utl_raw.cast_to_varchar2( utl_encode.base64_encode( c_salt ) ) || '"' ||
' encryptedVerifierHashInput="' || encryptedVerifierInput || '"' ||
' encryptedVerifierHashValue="' || encryptedVerifierValue || '"' ||
' encryptedKeyValue="' || encryptedKeyValue  || '"/></keyEncryptor></keyEncryptors></encryption>' ));
  end doEncryption;
  begin
    doEncryption( p_password, p_xlsx, t_EncryptedPackage, t_EncryptionInfo );
    --
    t_cf := utl_raw.copies( '00', 512 );
    dbms_lob.createtemporary( t_short_stream, true );
    t_root := add_dir_entry( 'Root Entry', c_DIR_Root );
    t_dummy := add_dir_entry( 'EncryptedPackage', c_DIR_Stream, t_root, t_EncryptedPackage );
    t_storage := add_dir_entry( 'DataSpaces', c_DIR_Storage, t_root, p_prefix => '0600' );
    t_dummy := add_dir_entry( 'Version', c_DIR_Stream, t_storage, c_version );
    t_dummy := add_dir_entry( 'DataSpaceMap', c_DIR_Stream, t_storage, c_DataSpaceMap );
    t_storage2 := add_dir_entry( 'DataSpaceInfo', c_DIR_Storage,t_storage );
    t_dummy := add_dir_entry( 'StrongEncryptionDataSpace', c_DIR_Stream, t_storage2, c_StrongEncryptionDataSpace );
    t_storage2 := add_dir_entry( 'TransformInfo', c_DIR_Storage, t_storage );
    t_storage2 := add_dir_entry( 'StrongEncryptionTransform', c_DIR_Storage, t_storage2 );
    t_dummy := add_dir_entry( 'Primary', c_DIR_Stream, t_storage2, c_Primary, p_prefix => '0600' );
    t_dummy := add_dir_entry( 'EncryptionInfo', c_DIR_Stream, t_root, t_EncryptionInfo );
    --
    dbms_lob.freetemporary( t_EncryptedPackage );
    --
    -- write the short sector stream
    dbms_lob.append( t_cf, t_short_stream );
    if mod( dbms_lob.getlength( t_short_stream ), t_ssz ) > 0
    then
      dbms_lob.writeappend( t_cf, t_ssz - mod( dbms_lob.getlength( t_short_stream ), t_ssz ), utl_raw.copies( '00', t_ssz ) );
    end if;
    t_dir( 0 ).len := dbms_lob.getlength( t_short_stream );
    t_dir( 0 ).first_sector := t_sat.count;
    for i in t_sat.count .. t_sat.count + trunc( ( dbms_lob.getlength( t_short_stream ) - 1 ) / t_ssz ) - 1
    loop
      t_sat( i ) := i + 1;
    end loop;
    t_sat( t_sat.count ) := c_End_Of_Chain_SecID;
    --
    -- write the ssat
    for i in 0 .. t_ssat.count - 1
    loop
      dbms_lob.writeappend( t_cf, 4, little_endian( t_ssat( i ) ) );
    end loop;
    if mod( t_ssat.count * 4, t_ssz ) > 0
    then
      dbms_lob.writeappend( t_cf, t_ssz - mod( t_ssat.count * 4, t_ssz ), utl_raw.copies( little_endian( c_Free_SecID ), t_ssz ) );
    end if;
    t_st_ssf := t_sat.count;
    for i in t_sat.count .. t_sat.count + trunc( ( t_ssat.count * 4 - 1 ) / t_ssz ) - 1
    loop
      t_sat( i ) := i + 1;
    end loop;
    t_sat( t_sat.count ) := c_End_Of_Chain_SecID;
    t_cnt_ssf := t_sat.count - t_st_ssf;
    --
    for i in 0 .. t_dir.last
    loop
      if t_dir( i ).childs.count = 1
      then
        t_dir( i ).root := t_dir( i ).childs( 0 );
        t_dir( t_dir( i ).childs( 0 ) ).colour := c_CLR_Black;
      elsif t_dir( i ).childs.count > 1
      then
        t_sorted := false;
        while not t_sorted
        loop
          t_sorted := true;
          for j in 0 .. t_dir( i ).childs.count - 2
          loop
            if is_less( t_dir( t_dir( i ).childs( j + 1 ) ), t_dir( t_dir( i ).childs( j ) ) )
            then
              t_tmp := t_dir( i ).childs( j ) ;
              t_dir( i ).childs( j ) := t_dir( i ).childs( j + 1 );
              t_dir( i ).childs( j + 1) := t_tmp;
              t_sorted := false;
            end if;
          end loop;
        end loop;
        --
        t_tmp := t_dir( i ).childs( 1 );
        t_dir( i ).root := t_tmp;
        t_dir( t_tmp ).left := t_dir( i ).childs( 0 );
        t_dir( t_tmp ).colour := c_CLR_Black;
        if t_dir( i ).childs.count > 2
        then
          t_dir( t_tmp ).right := t_dir( i ).childs( 2 );
          if t_dir( i ).childs.count > 3
          then
            t_dir( t_dir( i ).childs( 2 ) ).right := t_dir( i ).childs( 3 );
            t_dir( t_dir( i ).childs( 0 ) ).colour := c_CLR_Black;
            t_dir( t_dir( i ).childs( 2 ) ).colour := c_CLR_Black;
          end if;
        end if;
      end if;
    end loop;
    --
    -- write the dir tree
    for i in 0 .. t_dir.count - 1
    loop
      dbms_lob.writeappend( t_cf, 128
                          , utl_raw.concat( utl_raw.overlay( '00', t_dir( i ).rname, 64 )
                                          , little_endian( utl_raw.length( t_dir( i ).rname ) + 2, 2 )
                                          , t_dir( i ).tp_entry
                                          , t_dir( i ).colour
                                          , little_endian( t_dir( i ).left )
                                          , little_endian( t_dir( i ).right )
                                          , little_endian( t_dir( i ).root )
                                          , utl_raw.copies( '00', 36 )
                                          , little_endian( t_dir( i ).first_sector )
                                          , little_endian( t_dir( i ).len )
                                          , utl_raw.copies( '00', 4 )
                                          )
                          );
      end loop;
      if mod( t_dir.count * 128, t_ssz ) > 0
      then
        dbms_lob.writeappend( t_cf, t_ssz - mod( t_dir.count * 128, t_ssz ), utl_raw.copies( '00', t_ssz ) );
      end if;
      t_st_dir := t_sat.count;
      for i in t_st_dir .. t_st_dir + trunc( ( t_dir.count * 128 - 1 ) / t_ssz ) - 1
      loop
        t_sat( i ) := i + 1;
      end loop;
      t_sat( t_sat.count ) := c_End_Of_Chain_SecID;
      --
      -- write the sat
      t_tmp := floor( t_sat.count * 4 / t_ssz );
      for i in 0 .. t_tmp
      loop
        t_msat( t_msat.count ) := t_sat.count;
        t_sat( t_sat.count ) := c_SAT_SecID;
      end loop;
      if t_tmp != floor( t_sat.count * 4 / t_ssz )
      then
        t_msat( t_msat.count ) := t_sat.count;
        t_sat( t_sat.count ) := c_SAT_SecID;
      end if;
      for i in 0 .. t_sat.count - 1
      loop
        dbms_lob.writeappend( t_cf, 4, little_endian( t_sat( i ) ) );
      end loop;
      if mod( t_sat.count * 4, t_ssz ) > 0
      then
        dbms_lob.writeappend( t_cf, t_ssz - mod( t_sat.count * 4, t_ssz ), utl_raw.copies( little_endian( c_Free_SecID ), t_ssz ) );
      end if;
      t_header := utl_raw.concat( hextoraw( 'D0CF11E0A1B11AE1' )
                                , utl_raw.copies( '00', 16 )
                                , hextoraw( '3E000300' )
                                , hextoraw( 'FEFF' )
                                , little_endian( round( log( 2, t_ssz ) ), 2 )
                                , little_endian( round( log( 2, t_sssz ) ), 2 )
                                , utl_raw.copies( '00', 10 )
                                , little_endian( t_msat.count )
                                , little_endian( t_st_dir )
                                , utl_raw.copies( '00', 4 )
                                , little_endian( t_ss_cutoff )
                                , little_endian( t_st_ssf )
                                );
    t_header := utl_raw.concat( t_header
                              , little_endian( t_cnt_ssf )
                              , little_endian( c_End_Of_Chain_SecID )
                              , utl_raw.copies( '00', 4 )
                              );
    for i in 0 .. t_msat.count - 1
    loop
      t_header := utl_raw.concat( t_header
                                , little_endian( t_msat( i ) )
                                );
    end loop;
    t_header := utl_raw.concat( t_header
                              , utl_raw.copies( little_endian( c_Free_SecID ), 109 - t_msat.count )
                              );
    dbms_lob.copy( t_cf, t_header, 512, 1, 1 );
    dbms_lob.freetemporary( t_short_stream );
    return t_cf;
  end excel_encrypt;
$END

  --------------------------------------------------------------------------
  -- XLSX build procedures
  --
  -- Each procedure generates one part of the XLSX (Office Open XML) package
  -- and appends it to the ZIP archive BLOB.  They are called in sequence by
  -- finish().
  --------------------------------------------------------------------------

  -- Generates [Content_Types].xml — the master content type map.
  procedure build_content_types( p_excel in out nocopy blob )
  is
    l_xxx clob;
    l_s pls_integer;
  begin
    l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
    l_s := workbook.sheets.first;
    while l_s is not null
    loop
      l_xxx := l_xxx || ( '
<Override PartName="/xl/worksheets/sheet' || l_s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' );
      l_s := workbook.sheets.next( l_s );
    end loop;
    l_xxx := l_xxx || '
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    l_s := workbook.sheets.first;
    while l_s is not null
    loop
      if workbook.sheets( l_s ).comments.count > 0
      then
        l_xxx := l_xxx || ( '
<Override PartName="/xl/comments' || l_s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' );
      end if;
      if workbook.sheets( l_s ).drawings.count > 0
      then
        l_xxx := l_xxx || ( '
<Override ContentType="application/vnd.openxmlformats-officedocument.drawing+xml" PartName="/xl/drawings/drawing' || l_s || '.xml"/>' );
      end if;
      l_s := workbook.sheets.next( l_s );
    end loop;
    if workbook.images.count > 0
    then
      l_xxx := l_xxx || '
<Default ContentType="image/png" Extension="png"/>';
    end if;
    for i in 1 .. workbook.tables.count
    loop
      l_xxx := l_xxx || ( '
<Override PartName="/xl/tables/table' || to_char(i) || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' );
    end loop;
    l_xxx := l_xxx || '
</Types>';
    add1xml( p_excel, '[Content_Types].xml', l_xxx );
  end build_content_types;

  -- Generates xl/tables/table{n}.xml for each registered table.
  procedure build_tables( p_excel in out nocopy blob )
  is
    l_xxx clob;
    l_name  varchar2(32767);
    l_ref   varchar2(100);
    l_row   tp_cells;
    l_table tp_table;
    type tp_test is table of varchar2(32767);
    l_test  tp_test;
  begin
    for i in 1 .. workbook.tables.count
    loop
      l_table := workbook.tables( i );
      l_ref := '"' || alfan_col( l_table.column_start ) || l_table.row_start ||
               ':' || alfan_col( l_table.column_end ) || l_table.row_end || '"';
      l_xxx := ( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' ||
                 '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' ||
                 ' id="' || to_char( i ) || '"' ||
                 ' name="' || l_table.name || '"' ||
                 ' displayName="' || l_table.name || '"' ||
                 ' ref=' || l_ref ||
                 ' totalsRowShown="0">' ||
                 '<autoFilter ref=' || l_ref || '/>' ||
                 '<tableColumns count="' || to_char( 1 + l_table.column_end - l_table.column_start ) || '">'
               );
      l_test := tp_test();
      l_row := workbook.sheets( l_table.sheet ).rows( l_table.row_start );
      for j in l_table.column_start .. l_table.column_end
      loop
        l_name := workbook.str_ind( l_row( j ).value );
        if l_name member of l_test
        then
          raise_application_error( -20010, 'Table Header "' || l_name || '" appears multiple times in Table "' || l_table.name || '"' );
        end if;
        l_test := l_test multiset union tp_test( l_name );
        l_xxx := l_xxx || ( '<tableColumn id="' || to_char( j ) || '" name="' || l_name || '"/>' );
      end loop;
      l_xxx := l_xxx || ( '</tableColumns>' ||
                          '<tableStyleInfo name="' || l_table.style || '"' ||
                          ' showFirstColumn="0" showLastColumn="0"' ||
                          ' showRowStripes="1" showColumnStripes="0"' ||
                          '/></table>'
                        );
      add1xml( p_excel, 'xl/tables/table' || to_char(i) || '.xml', l_xxx );
    end loop;
  end build_tables;

  -- Generates docProps/core.xml (Dublin Core metadata: creator, timestamps).
  procedure build_core_properties( p_excel in out nocopy blob )
  is
    l_xxx clob;
  begin
    l_xxx := ( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>' || sys_context( 'userenv', 'os_user' ) || '</dc:creator>
<dc:description>Build by version:' || c_version || '</dc:description>
<cp:lastModifiedBy>' || sys_context( 'userenv', 'os_user' ) || '</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:modified>
</cp:coreProperties>' );
    add1xml( p_excel, 'docProps/core.xml', l_xxx );
  end build_core_properties;

  -- Generates docProps/app.xml (application metadata, sheet names).
  procedure build_app_properties( p_excel in out nocopy blob )
  is
    l_xxx clob;
    l_s pls_integer;
  begin
    l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<HeadingPairs>
<vt:vector size="2" baseType="variant">
<vt:variant>
<vt:lpstr>Worksheets</vt:lpstr>
</vt:variant>
<vt:variant>
<vt:i4>' || workbook.sheets.count || '</vt:i4>
</vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector size="' || workbook.sheets.count || '" baseType="lpstr">';
    l_s := workbook.sheets.first;
    while l_s is not null
    loop
      l_xxx := l_xxx || ( '
<vt:lpstr>' || workbook.sheets( l_s ).name || '</vt:lpstr>' );
      l_s := workbook.sheets.next( l_s );
    end loop;
    l_xxx := l_xxx || '</vt:vector>
</TitlesOfParts>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>14.0300</AppVersion>
</Properties>';
    add1xml( p_excel, 'docProps/app.xml', l_xxx );
  end build_app_properties;

  -- Generates _rels/.rels (top-level package relationships).
  procedure build_relationships( p_excel in out nocopy blob )
  is
    l_xxx clob;
  begin
    l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
    add1xml( p_excel, '_rels/.rels', l_xxx );
  end build_relationships;

  -- Generates xl/styles.xml (number formats, fonts, fills, borders, cell XFs).
  procedure build_styles( p_excel in out nocopy blob )
  is
    l_xxx clob;
  begin
    l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
    if workbook.numFmts.count > 0
    then
      l_xxx := l_xxx || ( '<numFmts count="' || workbook.numFmts.count || '">' );
      for n in 1 .. workbook.numFmts.count
      loop
        l_xxx := l_xxx || ( '<numFmt numFmtId="' || workbook.numFmts( n ).numFmtId || '" formatCode="' || workbook.numFmts( n ).formatCode || '"/>' );
      end loop;
      l_xxx := l_xxx || '</numFmts>';
    end if;
    l_xxx := l_xxx || ( '<fonts count="' || workbook.fonts.count || '" x14ac:knownFonts="1">' );
    for f in 0 .. workbook.fonts.count - 1
    loop
      l_xxx := l_xxx || ( '<font>' ||
        case when workbook.fonts( f ).bold then '<b/>' end ||
        case when workbook.fonts( f ).italic then '<i/>' end ||
        case when workbook.fonts( f ).underline then '<u/>' end ||
'<sz val="' || to_char( workbook.fonts( f ).fontsize, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )  || '"/>' ||
    case when workbook.fonts( f ).rgb is not null
      then add_rgb( workbook.fonts( f ).rgb )
      else '<color theme="' || workbook.fonts( f ).theme || '"/>'
    end ||
'<name val="' || workbook.fonts( f ).name || '"/>
<family val="' || workbook.fonts( f ).family || '"/>
<scheme val="none"/>
</font>' );
    end loop;
    l_xxx := l_xxx || ( '</fonts>
<fills count="' || workbook.fills.count || '">' );
    for f in 0 .. workbook.fills.count - 1
    loop
      l_xxx := l_xxx || ( '<fill><patternFill patternType="' || workbook.fills( f ).patternType || '">' ||
         add_rgb( workbook.fills( f ).fgRGB, 'fgColor' ) ||
         '</patternFill></fill>' );
    end loop;
    l_xxx := l_xxx || ( '</fills>
<borders count="' || workbook.borders.count || '">' );
    for b in 0 .. workbook.borders.count - 1
    loop
      l_xxx := l_xxx || ( '<border>' ||
         case when workbook.borders( b ).left   is null then '<left/>'   else '<left style="'   || workbook.borders( b ).left   || '">' || add_rgb( workbook.borders( b ).rgb ) || '</left>' end ||
         case when workbook.borders( b ).right  is null then '<right/>'  else '<right style="'  || workbook.borders( b ).right  || '">' || add_rgb( workbook.borders( b ).rgb ) || '</right>' end ||
         case when workbook.borders( b ).top    is null then '<top/>'    else '<top style="'    || workbook.borders( b ).top    || '">' || add_rgb( workbook.borders( b ).rgb ) || '</top>' end ||
         case when workbook.borders( b ).bottom is null then '<bottom/>' else '<bottom style="' || workbook.borders( b ).bottom || '">' || add_rgb( workbook.borders( b ).rgb ) || '</bottom>' end ||
         '</border>' );
    end loop;
    l_xxx := l_xxx || ( '</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="' || ( workbook.cellXfs.count + 1 ) || '">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' );
    for x in 1 .. workbook.cellXfs.count
    loop
      l_xxx := l_xxx || ( '<xf numFmtId="' || workbook.cellXfs( x ).numFmtId || '" fontId="' || workbook.cellXfs( x ).fontId || '" fillId="' || workbook.cellXfs( x ).fillId || '" borderId="' || workbook.cellXfs( x ).borderId || '">' );
      if (  workbook.cellXfs( x ).alignment.horizontal is not null
         or workbook.cellXfs( x ).alignment.vertical is not null
         or workbook.cellXfs( x ).alignment.wrapText
         or workbook.cellXfs( x ).alignment.rotation is not null
         )
      then
        l_xxx := l_xxx || ( '<alignment' ||
          case when workbook.cellXfs( x ).alignment.horizontal is not null then ' horizontal="' || workbook.cellXfs( x ).alignment.horizontal || '"' end ||
          case when workbook.cellXfs( x ).alignment.vertical is not null then ' vertical="' || workbook.cellXfs( x ).alignment.vertical || '"' end ||
          case when workbook.cellXfs( x ).alignment.rotation is not null then ' textRotation="' || round( workbook.cellXfs( x ).alignment.rotation ) || '"' end ||
          case when workbook.cellXfs( x ).alignment.wrapText then ' wrapText="true"' end || '/>' );
      end if;
      l_xxx := l_xxx || '</xf>';
    end loop;
    l_xxx := l_xxx || ( '</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
<extLst>
<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
</ext>
</extLst>
</styleSheet>' );
    add1xml( p_excel, 'xl/styles.xml', l_xxx );
  end build_styles;

  -- Generates xl/workbook.xml (sheet list, defined names, calc properties).
  procedure build_workbook( p_excel in out nocopy blob )
  is
    l_xxx clob;
    l_s pls_integer;
  begin
    l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
<workbookPr defaultThemeVersion="124226"/>
<bookViews>
<workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
</bookViews>
<sheets>';
    l_s := workbook.sheets.first;
    while l_s is not null
    loop
      l_xxx := l_xxx || ( '
<sheet name="' || workbook.sheets( l_s ).name || '" sheetId="' || l_s || '" r:id="rId' || ( 9 + l_s ) || '"/>' );
      l_s := workbook.sheets.next( l_s );
    end loop;
    l_xxx := l_xxx || '</sheets>';
    if workbook.defined_names.count > 0
    then
      l_xxx := l_xxx || '<definedNames>';
      for l_s in 1 .. workbook.defined_names.count
      loop
        l_xxx := l_xxx || ( '
<definedName name="' || workbook.defined_names( l_s ).name || '"' ||
            case when workbook.defined_names( l_s ).sheet is not null then ' localSheetId="' || to_char( workbook.defined_names( l_s ).sheet ) || '"' end ||
            '>' || workbook.defined_names( l_s ).ref || '</definedName>' );
      end loop;
      l_xxx := l_xxx || '</definedNames>';
    end if;
    l_xxx := l_xxx || '<calcPr calcId="144525"/></workbook>';
    add1xml( p_excel, 'xl/workbook.xml', l_xxx );
  end build_workbook;

  -- Generates xl/theme/theme1.xml.
  -- The theme XML is declared as a local constant to keep the large static
  -- string out of the package-level declaration section.
  procedure build_theme( p_excel in out nocopy blob )
  is
    c_THEME_XML constant clob := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1>
<a:sysClr val="windowText" lastClr="000000"/>
</a:dk1>
<a:lt1>
<a:sysClr val="window" lastClr="FFFFFF"/>
</a:lt1>
<a:dk2>
<a:srgbClr val="1F497D"/>
</a:dk2>
<a:lt2>
<a:srgbClr val="EEECE1"/>
</a:lt2>
<a:accent1>
<a:srgbClr val="4F81BD"/>
</a:accent1>
<a:accent2>
<a:srgbClr val="C0504D"/>
</a:accent2>
<a:accent3>
<a:srgbClr val="9BBB59"/>
</a:accent3>
<a:accent4>
<a:srgbClr val="8064A2"/>
</a:accent4>
<a:accent5>
<a:srgbClr val="4BACC6"/>
</a:accent5>
<a:accent6>
<a:srgbClr val="F79646"/>
</a:accent6>
<a:hlink>
<a:srgbClr val="0000FF"/>
</a:hlink>
<a:folHlink>
<a:srgbClr val="800080"/>
</a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont>
<a:latin typeface="Cambria"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Times New Roman"/>
<a:font script="Hebr" typeface="Times New Roman"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="MoolBoran"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Times New Roman"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="Calibri"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Arial"/>
<a:font script="Hebr" typeface="Arial"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="DaunPenh"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Arial"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="50000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="35000">
<a:schemeClr val="phClr">
<a:tint val="37000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:tint val="15000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="1"/>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:shade val="51000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="80000">
<a:schemeClr val="phClr">
<a:shade val="93000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="94000"/>
<a:satMod val="135000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="0"/>
</a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr">
<a:shade val="95000"/>
<a:satMod val="105000"/>
</a:schemeClr>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="38000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
<a:scene3d>
<a:camera prst="orthographicFront">
<a:rot lat="0" lon="0" rev="0"/>
</a:camera>
<a:lightRig rig="threePt" dir="t">
<a:rot lat="0" lon="0" rev="1200000"/>
</a:lightRig>
</a:scene3d>
<a:sp3d>
<a:bevelT w="63500" h="25400"/>
</a:sp3d>
</a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="40000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="40000">
<a:schemeClr val="phClr">
<a:tint val="45000"/>
<a:shade val="99000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="20000"/>
<a:satMod val="255000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
</a:path>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="80000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="30000"/>
<a:satMod val="200000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
</a:path>
</a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>';

  begin
    add1xml( p_excel, 'xl/theme/theme1.xml', c_THEME_XML );
  end build_theme;

  -- Generates all worksheet XML files:
  --   xl/worksheets/sheet{n}.xml   — cell data, panes, autofilter, validations
  --   xl/comments{n}.xml          — cell comments
  --   xl/drawings/vmlDrawing{n}.vml — comment shapes (VML)
  --   xl/drawings/drawing{n}.xml  — embedded images (DrawingML)
  --   + associated .rels files
  procedure build_worksheets( p_excel in out nocopy blob )
  is
    l_xxx clob;
    l_yyy blob;
    l_c number;
    l_h number;
    l_w number;
    l_cw number;
    l_s pls_integer;
    l_row_ind pls_integer;
    l_col_min pls_integer;
    l_col_max pls_integer;
    l_col_ind pls_integer;
  begin
    l_s := workbook.sheets.first;
    while l_s is not null
    loop
      l_col_min := 16384;
      l_col_max := 1;
      l_row_ind := workbook.sheets( l_s ).rows.first;
      while l_row_ind is not null
      loop
        l_col_min := least( l_col_min, nvl( workbook.sheets( l_s ).rows( l_row_ind ).first, l_col_min ) );
        l_col_max := greatest( l_col_max, nvl( workbook.sheets( l_s ).rows( l_row_ind ).last, l_col_max ) );
        l_row_ind := workbook.sheets( l_s ).rows.next( l_row_ind );
      end loop;
      if l_col_min = 16384
      then -- no "cell" in sheet, only images for instance see https://github.com/antonscheffer/as_xlsx/issues/23
        l_col_min := l_col_max;
      end if;
      addtxt2utf8blob_init( l_yyy );
      addtxt2utf8blob( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' ||
case when workbook.sheets( l_s ).tabcolor is not null then '<sheetPr>' || add_rgb( workbook.sheets( l_s ).tabcolor, 'tabColor' ) || '</sheetPr>' end ||
'<dimension ref="' || alfan_col( l_col_min ) || workbook.sheets( l_s ).rows.first || ':' || alfan_col( l_col_max ) || workbook.sheets( l_s ).rows.last || '"/>
<sheetViews>
<sheetView' ||
  case when workbook.sheets( l_s ).grid_color_idx is not null then ' defaultGridColor="0" colorId="' || to_char( workbook.sheets( l_s ).grid_color_idx ) || '"' end ||
  case when not workbook.sheets( l_s ).show_gridlines then ' showGridLines="0"' end ||
  case when not workbook.sheets( l_s ).show_headers then ' showRowColHeaders="0"' end ||
  case when l_s = 1 then ' tabSelected="1"' end || ' workbookViewId="0">'
                     , l_yyy
                     );
      if workbook.sheets( l_s ).freeze_rows > 0 and workbook.sheets( l_s ).freeze_cols > 0
      then
        addtxt2utf8blob( '<pane xSplit="' || workbook.sheets( l_s ).freeze_cols || '" '
                          || 'ySplit="' || workbook.sheets( l_s ).freeze_rows || '" '
                          || 'topLeftCell="' || alfan_col( workbook.sheets( l_s ).freeze_cols + 1 ) || ( workbook.sheets( l_s ).freeze_rows + 1 ) || '" '
                          || 'activePane="bottomLeft" state="frozen"/>'
                       , l_yyy
                       );
      else
        if workbook.sheets( l_s ).freeze_rows > 0
        then
          addtxt2utf8blob( '<pane ySplit="' || workbook.sheets( l_s ).freeze_rows || '" topLeftCell="A' || ( workbook.sheets( l_s ).freeze_rows + 1 ) || '" activePane="bottomLeft" state="frozen"/>'
                         , l_yyy
                         );
        end if;
        if workbook.sheets( l_s ).freeze_cols > 0
        then
          addtxt2utf8blob( '<pane xSplit="' || workbook.sheets( l_s ).freeze_cols || '" topLeftCell="' || alfan_col( workbook.sheets( l_s ).freeze_cols + 1 ) || '1" activePane="bottomLeft" state="frozen"/>'
                         , l_yyy
                         );
        end if;
      end if;
      addtxt2utf8blob( '</sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
                     , l_yyy
                     );
      if workbook.sheets( l_s ).widths.count > 0
      then
        addtxt2utf8blob( '<cols>', l_yyy );
        l_col_ind := workbook.sheets( l_s ).widths.first;
        while l_col_ind is not null
        loop
          addtxt2utf8blob( '<col min="' || l_col_ind || '" max="' || l_col_ind || '" width="' || to_char( workbook.sheets( l_s ).widths( l_col_ind ), 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" customWidth="1"/>', l_yyy );
          l_col_ind := workbook.sheets( l_s ).widths.next( l_col_ind );
        end loop;
        addtxt2utf8blob( '</cols>', l_yyy );
      end if;
      addtxt2utf8blob( '<sheetData>', l_yyy );
      l_row_ind := workbook.sheets( l_s ).rows.first;
      while l_row_ind is not null
      loop
        if workbook.sheets( l_s ).row_fmts.exists( l_row_ind ) and workbook.sheets( l_s ).row_fmts( l_row_ind ).height is not null
        then
          addtxt2utf8blob( '<row r="' || l_row_ind || '" spans="' || l_col_min || ':' || l_col_max || '" customHeight="1" ht="'
                         || to_char( workbook.sheets( l_s ).row_fmts( l_row_ind ).height, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || '" >', l_yyy );
        else
          addtxt2utf8blob( '<row r="' || l_row_ind || '" spans="' || l_col_min || ':' || l_col_max || '">', l_yyy );
        end if;
        l_col_ind := workbook.sheets( l_s ).rows( l_row_ind ).first;
        while l_col_ind is not null
        loop
          addtxt2utf8blob( '<c r="' || alfan_col( l_col_ind ) || l_row_ind || '"'
                 || ' ' || workbook.sheets( l_s ).rows( l_row_ind )( l_col_ind ).style
                 || '>' || workbook.sheets( l_s ).rows( l_row_ind )( l_col_ind ).formula || '<v>'
                 || to_char( workbook.sheets( l_s ).rows( l_row_ind )( l_col_ind ).value, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )
                 || '</v></c>', l_yyy );
          l_col_ind := workbook.sheets( l_s ).rows( l_row_ind ).next( l_col_ind );
        end loop;
        addtxt2utf8blob( '</row>', l_yyy );
        l_row_ind := workbook.sheets( l_s ).rows.next( l_row_ind );
      end loop;
      addtxt2utf8blob( '</sheetData>', l_yyy );
      for a in 1 ..  workbook.sheets( l_s ).autofilters.count
      loop
        addtxt2utf8blob( '<autoFilter ref="' ||
            alfan_col( nvl( workbook.sheets( l_s ).autofilters( a ).column_start, l_col_min ) ) ||
            nvl( workbook.sheets( l_s ).autofilters( a ).row_start, workbook.sheets( l_s ).rows.first ) || ':' ||
            alfan_col( coalesce( workbook.sheets( l_s ).autofilters( a ).column_end, workbook.sheets( l_s ).autofilters( a ).column_start, l_col_max ) ) ||
            nvl( workbook.sheets( l_s ).autofilters( a ).row_end, workbook.sheets( l_s ).rows.last ) || '"/>', l_yyy );
      end loop;
      if workbook.sheets( l_s ).mergecells.count > 0
      then
        addtxt2utf8blob( '<mergeCells count="' || to_char( workbook.sheets( l_s ).mergecells.count ) || '">', l_yyy );
        for m in 1 ..  workbook.sheets( l_s ).mergecells.count
        loop
          addtxt2utf8blob( '<mergeCell ref="' || workbook.sheets( l_s ).mergecells( m ) || '"/>', l_yyy );
        end loop;
        addtxt2utf8blob( '</mergeCells>', l_yyy );
      end if;
--
      if workbook.sheets( l_s ).validations.count > 0
      then
        addtxt2utf8blob( '<dataValidations count="' || to_char( workbook.sheets( l_s ).validations.count ) || '">', l_yyy );
        for m in 1 ..  workbook.sheets( l_s ).validations.count
        loop
          addtxt2utf8blob( '<dataValidation' ||
              ' type="' || workbook.sheets( l_s ).validations( m ).type || '"' ||
              ' errorStyle="' || workbook.sheets( l_s ).validations( m ).errorstyle || '"' ||
              ' allowBlank="' || case when nvl( workbook.sheets( l_s ).validations( m ).allowBlank, true ) then '1' else '0' end || '"' ||
              ' sqref="' || workbook.sheets( l_s ).validations( m ).sqref || '"', l_yyy );
          if workbook.sheets( l_s ).validations( m ).prompt is not null
          then
            addtxt2utf8blob( ' showInputMessage="1" prompt="' || workbook.sheets( l_s ).validations( m ).prompt || '"', l_yyy );
            if workbook.sheets( l_s ).validations( m ).title is not null
            then
              addtxt2utf8blob( ' promptTitle="' || workbook.sheets( l_s ).validations( m ).title || '"', l_yyy );
            end if;
          end if;
          if workbook.sheets( l_s ).validations( m ).showerrormessage
          then
            addtxt2utf8blob( ' showErrorMessage="1"', l_yyy );
            if workbook.sheets( l_s ).validations( m ).error_title is not null
            then
              addtxt2utf8blob( ' errorTitle="' || workbook.sheets( l_s ).validations( m ).error_title || '"', l_yyy );
            end if;
            if workbook.sheets( l_s ).validations( m ).error_txt is not null
            then
              addtxt2utf8blob( ' error="' || workbook.sheets( l_s ).validations( m ).error_txt || '"', l_yyy );
            end if;
          end if;
          addtxt2utf8blob( '>', l_yyy );
          if workbook.sheets( l_s ).validations( m ).formula1 is not null
          then
            addtxt2utf8blob( '<formula1>' || workbook.sheets( l_s ).validations( m ).formula1 || '</formula1>', l_yyy );
          end if;
          if workbook.sheets( l_s ).validations( m ).formula2 is not null
          then
            addtxt2utf8blob( '<formula2>' || workbook.sheets( l_s ).validations( m ).formula2 || '</formula2>', l_yyy );
          end if;
          addtxt2utf8blob( '</dataValidation>', l_yyy );
        end loop;
        addtxt2utf8blob( '</dataValidations>', l_yyy );
      end if;
--
      if workbook.sheets( l_s ).hyperlinks.count > 0
      then
        addtxt2utf8blob( '<hyperlinks>', l_yyy );
        for h in 1 ..  workbook.sheets( l_s ).hyperlinks.count
        loop
          addtxt2utf8blob( '<hyperlink ref="' || workbook.sheets( l_s ).hyperlinks( h ).cell || '"', l_yyy );
          if workbook.sheets( l_s ).hyperlinks( h ).url is null
          then
            addtxt2utf8blob( ' location="' || workbook.sheets( l_s ).hyperlinks( h ).location || '"', l_yyy );
          else
            addtxt2utf8blob( ' r:id="rId' || ( 3 + h ) || '"', l_yyy );
          end if;
          if workbook.sheets( l_s ).hyperlinks( h ).tooltip is not null
          then
            addtxt2utf8blob( ' tooltip="' || workbook.sheets( l_s ).hyperlinks( h ).tooltip || '"/>', l_yyy );
          else
            addtxt2utf8blob( '/>', l_yyy );
          end if;
        end loop;
        addtxt2utf8blob( '</hyperlinks>', l_yyy );
      end if;
      addtxt2utf8blob( '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>', l_yyy );
      if workbook.sheets( l_s ).drawings.count > 0
      then
        addtxt2utf8blob( '<drawing r:id="rId3"/>', l_yyy );
      end if;
      if workbook.sheets( l_s ).comments.count > 0
      then
        addtxt2utf8blob( '<legacyDrawing r:id="rId1"/>', l_yyy );
      end if;
      --
      declare
        l_cnt pls_integer := 0;
      begin
        for i in 1 .. workbook.tables.count
        loop
          if workbook.tables( i ).sheet = l_s
          then
            l_cnt := l_cnt + 1;
          end if;
        end loop;
        if l_cnt > 0
        then
          addtxt2utf8blob( '<tableParts count="' || to_char( l_cnt ) || '">', l_yyy) ;
          for i in 1 .. workbook.tables.count
          loop
            if workbook.tables( i ).sheet = l_s
            then
              addtxt2utf8blob( '<tablePart r:id="rId' || to_char( 10000 + i ) || '"/>', l_yyy );
            end if;
          end loop;
          addtxt2utf8blob( '</tableParts>', l_yyy );
        end if;
      end;
      --
      addtxt2utf8blob( '</worksheet>', l_yyy );
      addtxt2utf8blob_finish( l_yyy );
      add1file( p_excel, 'xl/worksheets/sheet' || l_s || '.xml', l_yyy );
      --
      if workbook.images.count > 0
      then
        l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        for i in 1 .. workbook.images.count
        loop
          l_xxx := l_xxx || ( '<Relationship Id="rId'
                         || i || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' || 'image' || i || '.png'
                         || '"/>' );
        end loop;
        l_xxx := l_xxx || '</Relationships>';
        add1xml( p_excel, 'xl/drawings/_rels/drawing' || l_s || '.xml.rels', l_xxx );
      end if;
      --
      if (  workbook.sheets( l_s ).hyperlinks.count > 0
         or workbook.sheets( l_s ).comments.count > 0
         or workbook.sheets( l_s ).drawings.count > 0
         or workbook.tables.count > 0
         )
      then
        l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        if workbook.sheets( l_s ).comments.count > 0
        then
          l_xxx := l_xxx || ( '<Relationship Id="rId2" Target="../comments' || l_s || '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"/>' );
          l_xxx := l_xxx || ( '<Relationship Id="rId1" Target="../drawings/vmlDrawing' || l_s || '.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/>' );
        end if;
        if workbook.sheets( l_s ).drawings.count > 0
        then
          l_xxx := l_xxx || ( '<Relationship Id="rId3" Target="../drawings/drawing' || l_s || '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>' );
        end if;
        for h in 1 ..  workbook.sheets( l_s ).hyperlinks.count
        loop
          if workbook.sheets( l_s ).hyperlinks( h ).url is not null
          then
            l_xxx := l_xxx || ( '<Relationship Id="rId' || ( 3 + h ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' || workbook.sheets( l_s ).hyperlinks( h ).url || '" TargetMode="External"/>' );
          end if;
        end loop;
        for i in 1 .. workbook.tables.count
        loop
          if workbook.tables( i ).sheet = l_s
          then
            l_xxx := l_xxx ||  ( '<Relationship Id="rId' || to_char( 10000 + i ) || '"' ||
                                 ' Target="../tables/table' || to_char( i ) || '.xml"' ||
                                 ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"/>' );
          end if;
        end loop;
        l_xxx := l_xxx || '</Relationships>';
        add1xml( p_excel, 'xl/worksheets/_rels/sheet' || l_s || '.xml.rels', l_xxx );
        --
        if workbook.sheets( l_s ).drawings.count > 0
        then
          l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
          for i in 1 .. workbook.sheets( l_s ).drawings.count
          loop
            l_xxx := l_xxx || finish_drawing( workbook.sheets( l_s ).drawings( i ), i, l_s );
          end loop;
          l_xxx := l_xxx || '</xdr:wsDr>';
          add1xml( p_excel, 'xl/drawings/drawing' || l_s || '.xml', l_xxx );
        end if;
--
        if workbook.sheets( l_s ).comments.count > 0
        then
          declare
            cnt pls_integer;
            author_ind tp_author;
          begin
            authors.delete;
            for c in 1 .. workbook.sheets( l_s ).comments.count
            loop
              authors( nvl( workbook.sheets( l_s ).comments( c ).author, ' ' ) ) := 0;
            end loop;
            l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>';
            cnt := 0;
            author_ind := authors.first;
            while author_ind is not null or authors.next( author_ind ) is not null
            loop
              authors( author_ind ) := cnt;
              l_xxx := l_xxx || ( '<author>' || author_ind || '</author>' );
              cnt := cnt + 1;
              author_ind := authors.next( author_ind );
            end loop;
          end;
          l_xxx := l_xxx || '</authors><commentList>';
          for c in 1 .. workbook.sheets( l_s ).comments.count
          loop
            l_xxx := l_xxx || ( '<comment ref="' || alfan_col( workbook.sheets( l_s ).comments( c ).column ) ||
               to_char( workbook.sheets( l_s ).comments( c ).row ) || '"' ||
               ' authorId="' || authors( nvl( workbook.sheets( l_s ).comments( c ).author, ' ' ) ) || '"><text>' );
            if workbook.sheets( l_s ).comments( c ).author is not null
            then
              l_xxx := l_xxx || ( '<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
                 workbook.sheets( l_s ).comments( c ).author || ':</t></r>' );
            end if;
            l_xxx := l_xxx || ( '<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
               case when workbook.sheets( l_s ).comments( c ).author is not null then '
' end || workbook.sheets( l_s ).comments( c ).text || '</t></r></text></comment>' );
          end loop;
          l_xxx := l_xxx || '</commentList></comments>';
          add1xml( p_excel, 'xl/comments' || l_s || '.xml', l_xxx );
          l_xxx := '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
          for c in 1 .. workbook.sheets( l_s ).comments.count
          loop
            l_xxx := l_xxx || ( '<v:shape id="_x0000_s' || to_char( c ) || '" type="#_x0000_t202"
style="position:absolute;margin-left:35.25pt;margin-top:3pt;z-index:' || to_char( c ) || ';visibility:hidden;" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' );
            l_w := workbook.sheets( l_s ).comments( c ).width;
            l_c := 1;
            loop
              if workbook.sheets( l_s ).widths.exists( workbook.sheets( l_s ).comments( c ).column + l_c )
              then
                l_cw := 256 * workbook.sheets( l_s ).widths( workbook.sheets( l_s ).comments( c ).column + l_c );
                l_cw := trunc( ( l_cw + 18 ) / 256 * 7); -- assume default 11 point Calibri
              else
                l_cw := 64;
              end if;
              exit when l_w < l_cw;
              l_c := l_c + 1;
              l_w := l_w - l_cw;
            end loop;
            l_h := workbook.sheets( l_s ).comments( c ).height;
            l_xxx := l_xxx || ( '<x:Anchor>' || workbook.sheets( l_s ).comments( c ).column || ',15,' ||
                       workbook.sheets( l_s ).comments( c ).row || ',30,' ||
                       ( workbook.sheets( l_s ).comments( c ).column + l_c - 1 ) || ',' || round( l_w ) || ',' ||
                       ( workbook.sheets( l_s ).comments( c ).row + 1 + trunc( l_h / 20 ) ) || ',' || mod( l_h, 20 ) || '</x:Anchor>' );
            l_xxx := l_xxx || ( '<x:AutoFill>False</x:AutoFill><x:Row>' ||
              ( workbook.sheets( l_s ).comments( c ).row - 1 ) || '</x:Row><x:Column>' ||
              ( workbook.sheets( l_s ).comments( c ).column - 1 ) || '</x:Column></x:ClientData></v:shape>' );
          end loop;
          l_xxx := l_xxx || '</xml>';
          add1xml( p_excel, 'xl/drawings/vmlDrawing' || l_s || '.vml', l_xxx );
        end if;
      end if;
--
      l_s := workbook.sheets.next( l_s );
    end loop;
  end build_worksheets;

  -- Generates xl/_rels/workbook.xml.rels (shared strings, styles, theme, sheets).
  procedure build_workbook_rels( p_excel in out nocopy blob )
  is
    l_xxx clob;
    l_s pls_integer;
  begin
    l_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
    l_s := workbook.sheets.first;
    while l_s is not null
    loop
      l_xxx := l_xxx || ( '
<Relationship Id="rId' || ( 9 + l_s ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' || l_s || '.xml"/>' );
      l_s := workbook.sheets.next( l_s );
    end loop;
    l_xxx := l_xxx || '</Relationships>';
    add1xml( p_excel, 'xl/_rels/workbook.xml.rels', l_xxx );
  end build_workbook_rels;

  -- Adds all image BLOBs to the ZIP as xl/media/image{n}.png.
  procedure build_images( p_excel in out nocopy blob )
  is
  begin
    for i in 1 .. workbook.images.count
    loop
      add1file( p_excel, 'xl/media/image' || i || '.png', workbook.images(i).img );
    end loop;
  end build_images;

  -- Generates xl/sharedStrings.xml from the workbook string table.
  procedure build_shared_strings( p_excel in out nocopy blob )
  is
    l_yyy blob;
  begin
    addtxt2utf8blob_init( l_yyy );
    addtxt2utf8blob( '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || workbook.str_cnt || '" uniqueCount="' || workbook.strings.count || '">'
                  , l_yyy
                  );
    for i in 0 .. workbook.str_ind.count - 1
    loop
      addtxt2utf8blob( '<si><t xml:space="preserve">' || dbms_xmlgen.convert( substr( workbook.str_ind( i ), 1, 32000 ) ) || '</t></si>', l_yyy );
    end loop;
    addtxt2utf8blob( '</sst>', l_yyy );
    addtxt2utf8blob_finish( l_yyy );
    add1file( p_excel, 'xl/sharedStrings.xml', l_yyy );
  end build_shared_strings;

  --------------------------------------------------------------------------
  -- finish — Main entry point for XLSX generation
  --
  -- Orchestrates all build_* procedures in the correct order, finalises
  -- the ZIP archive, clears workbook state, and optionally encrypts
  -- the result with a password.
  --------------------------------------------------------------------------
  function finish( p_password varchar2 := null )
  return blob
  is
    l_excel blob;
  begin
    dbms_lob.createtemporary( l_excel, true );
    build_content_types( l_excel );
    build_tables( l_excel );
    build_core_properties( l_excel );
    build_app_properties( l_excel );
    build_relationships( l_excel );
    build_styles( l_excel );
    build_workbook( l_excel );
    build_theme( l_excel );
    build_worksheets( l_excel );
    build_workbook_rels( l_excel );
    build_images( l_excel );
    build_shared_strings( l_excel );
    finish_zip( l_excel );
    clear_workbook;
$IF pck_as_xlsx.use_dbms_crypto
$THEN
    if p_password is not null
    then
      return excel_encrypt( l_excel, p_password );
    else
      return l_excel;
    end if;
$ELSE
    return l_excel;
$END
  end finish;

  --------------------------------------------------------------------------
  -- Query to sheet
  --
  -- Populates a worksheet directly from a SQL query or SYS_REFCURSOR.
  -- Handles NUMBER, DATE, and VARCHAR2 column types.  Supports optional
  -- column headers, title row, autofilter, and table formatting.
  --------------------------------------------------------------------------

  -- Core implementation: fetches rows from an already-opened DBMS_SQL cursor.
  function query2sheet
    ( p_c              in out integer
    , p_column_headers boolean
    , p_sheet          pls_integer
    , p_UseXf          boolean
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2
    , p_title_xfid     pls_integer
    , p_col            pls_integer
    , p_row            pls_integer
    , p_autofilter     boolean
    , p_table_style    varchar2
    )
  return number
  is
    l_null varchar2(1);
    l_sheet pls_integer;
    l_col_cnt integer;
    l_desc_tab dbms_sql.desc_tab2;
    l_d_tab dbms_sql.date_table;
    l_n_tab dbms_sql.number_table;
    l_v_tab dbms_sql.varchar2_table;
    l_bulk_size pls_integer := 200;
    l_r integer;
    l_col     pls_integer;
    l_cur_row pls_integer;
    l_useXf boolean := g_useXf;
    type tp_XfIds is table of varchar2(50) index by pls_integer;
    l_XfIds tp_XfIds;
    l_null_number number;
    l_rows number;
    l_horizontal varchar2(3999);
    l_right       boolean;
    l_center_cont boolean;
  begin
    if p_sheet is null
    then
      new_sheet;
    end if;
    l_sheet := coalesce( p_sheet, workbook.sheets.count );
    setUseXf( true );
    l_col := coalesce( p_col, 1 ) - 1;
    l_cur_row := coalesce( p_row, 1 );
    if p_title is not null
    then
      l_cur_row := l_cur_row + 1;
      if p_title_xfid is not null and workbook.cellXfs.exists( p_title_xfid )
      then
        l_horizontal := lower( workbook.cellXfs( p_title_xfid ).alignment.horizontal );
        l_right := l_horizontal = 'right';
        l_center_cont := l_horizontal = 'centercontinuous';
      end if;
      l_right := nvl( l_right, false );
      l_center_cont := nvl( l_center_cont, false );
    end if;
    dbms_sql.describe_columns2( p_c, l_col_cnt, l_desc_tab );
    for c in 1 .. l_col_cnt
    loop
      if p_title is not null
      then
        if ( c = 1 and not l_right ) or ( c = l_col_cnt and l_right )
        then
          cell( l_col + c, l_cur_row - 1, p_title, p_sheet => l_sheet );
          if p_title_xfid is not null
          then
            workbook.sheets( l_sheet ).rows( l_cur_row - 1 )( l_col + c ).style := 't="s" s="' || p_title_xfid || '"';
          end if;
        elsif l_center_cont
        then
          cell( l_col + c, l_cur_row - 1, l_null_number, p_sheet => l_sheet );
          workbook.sheets( l_sheet ).rows( l_cur_row - 1 )( l_col + c ).style := 's="' || p_title_xfid || '"';
        end if;
      end if;
      if p_column_headers
      then
        cell( l_col + c, l_cur_row, l_desc_tab( c ).col_name, p_sheet => l_sheet );
      end if;
      case
        when l_desc_tab( c ).col_type in ( 2, 100, 101 )
        then
          dbms_sql.define_array( p_c, c, l_n_tab, l_bulk_size, 1 );
        when l_desc_tab( c ).col_type in ( 12, 178, 179, 180, 181, 231 )
        then
          dbms_sql.define_array( p_c, c, l_d_tab, l_bulk_size, 1 );
          l_XfIds(c) := get_XfId( l_sheet, l_col + c, l_cur_row, get_numFmt( coalesce( p_date_format, 'dd/mm/yyyy' ) ), null, null );
        when l_desc_tab( c ).col_type in ( 1, 8, 9, 96, 112 )
        then
          dbms_sql.define_array( p_c, c, l_v_tab, l_bulk_size, 1 );
        else
          null;
      end case;
    end loop;
    --
    setUseXf( p_UseXf );
    if p_column_headers
    then
      l_cur_row := l_cur_row + 1;
    end if;
    --
    l_rows := 0;
    loop
      l_r := dbms_sql.fetch_rows( p_c );
      if l_r > 0
      then
        for c in 1 .. l_col_cnt
        loop
          case
            when l_desc_tab( c ).col_type in ( 2, 100, 101 )
            then
              dbms_sql.column_value( p_c, c, l_n_tab );
              for i in 0 .. l_r - 1
              loop
                if l_n_tab( i + l_n_tab.first ) is not null
                then
                  cell( l_col + c, l_cur_row + i, l_n_tab( i + l_n_tab.first ), p_sheet => l_sheet );
                else
                  cell( l_col + c, l_cur_row + i, l_null_number, p_sheet => l_sheet );
                end if;
              end loop;
              l_n_tab.delete;
            when l_desc_tab( c ).col_type in ( 12, 178, 179, 180, 181, 231 )
            then
              dbms_sql.column_value( p_c, c, l_d_tab );
              for i in 0 .. l_r - 1
              loop
                if l_d_tab( i + l_d_tab.first ) is not null
                then
                  if g_useXf
                  then
                    cell( l_col + c, l_cur_row + i, l_d_tab( i + l_d_tab.first ), p_sheet => l_sheet );
                  else
                    query_date_cell( l_col + c, l_cur_row + i, l_d_tab( i + l_d_tab.first ), l_sheet, l_XfIds(c) );
                  end if;
                else
                  cell( l_col + c, l_cur_row + i, l_null_number, p_sheet => l_sheet );
                end if;
              end loop;
              l_d_tab.delete;
            when l_desc_tab( c ).col_type in ( 1, 8, 9, 96, 112 )
            then
              dbms_sql.column_value( p_c, c, l_v_tab );
              for i in 0 .. l_r - 1
              loop
                if l_v_tab( i + l_v_tab.first ) is not null
                then
                  cell( l_col + c, l_cur_row + i, l_v_tab( i + l_v_tab.first ), p_sheet => l_sheet );
                else
                  cell( l_col + c, l_cur_row + i, l_null_number, p_sheet => l_sheet );
                end if;
              end loop;
              l_v_tab.delete;
            else
              null;
          end case;
        end loop;
      end if;
      l_rows := l_rows + l_r;
      l_cur_row := l_cur_row + l_r;
      exit when l_r != l_bulk_size;
    end loop;
    dbms_sql.close_cursor( p_c );
    --
    if p_autofilter and p_table_style is null and p_column_headers
    then
      set_autofilter
        ( p_column_start => l_col + 1
        , p_column_end   => l_col + l_col_cnt
        , p_row_start    => coalesce( p_row, 1 ) + case when p_title is null then 0 else 1 end
        , p_row_end      => l_cur_row - 1
        , p_sheet        => l_sheet
        );
    end if;
    --
    if p_table_style is not null and p_column_headers
    then
      set_table
        ( p_column_start => l_col + 1
        , p_column_end   => l_col + l_col_cnt
        , p_row_start    => coalesce( p_row, 1 ) + case when p_title is null then 0 else 1 end
        , p_row_end      => l_cur_row - 1
        , p_style        => p_table_style
        , p_sheet        => l_sheet
        );
    end if;
    --
    setUseXf( l_useXf );
    return l_rows;
  exception
    when others
    then
      if dbms_sql.is_open( p_c )
      then
        dbms_sql.close_cursor( p_c );
      end if;
      setUseXf( l_useXf );
      return null;
  end query2sheet;

  -- Convenience overload: accepts a SQL string, opens a cursor, and delegates.
  function query2sheet
    ( p_sql            varchar2
    , p_column_headers boolean     := true
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  return number
  is
    l_c integer;
    l_r integer;
  begin
    l_c := dbms_sql.open_cursor;
    dbms_sql.parse( l_c, p_sql, dbms_sql.native );
    l_r := dbms_sql.execute( l_c );
    return query2sheet
             ( p_c              => l_c
             , p_column_headers => p_column_headers
             , p_sheet          => p_sheet
             , p_UseXf          => p_UseXf
             , p_date_format    => p_date_format
             , p_title          => p_title
             , p_title_xfid     => p_title_xfid
             , p_col            => p_col
             , p_row            => p_row
             , p_autofilter     => p_autofilter
             , p_table_style    => p_table_style
             );
  end;

  -- Convenience overload: accepts a SYS_REFCURSOR, converts to DBMS_SQL, and delegates.
  function query2sheet
    ( p_rc             in out sys_refcursor
    , p_column_headers boolean     := true
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  return number
  is
    l_c integer;
    l_r integer;
  begin
    l_c := dbms_sql.to_cursor_number( p_rc );
    return query2sheet
             ( p_c              => l_c
             , p_column_headers => p_column_headers
             , p_sheet          => p_sheet
             , p_UseXf          => p_UseXf
             , p_date_format    => p_date_format
             , p_title          => p_title
             , p_title_xfid     => p_title_xfid
             , p_col            => p_col
             , p_row            => p_row
             , p_autofilter     => p_autofilter
             , p_table_style    => p_table_style
             );
  end;

  -- Procedure overload (SQL string): ignores the return value.
  procedure query2sheet
    ( p_sql            varchar2
    , p_column_headers boolean     := true
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  is
    l_dummy number;
  begin
    l_dummy := query2sheet
                 ( p_sql            => p_sql
                 , p_column_headers => p_column_headers
                 , p_sheet          => p_sheet
                 , p_UseXf          => p_UseXf
                 , p_date_format    => p_date_format
                 , p_title          => p_title
                 , p_title_xfid     => p_title_xfid
                 , p_col            => p_col
                 , p_row            => p_row
                 , p_autofilter     => p_autofilter
                 , p_table_style    => p_table_style
                 );
  end;

  -- Procedure overload (SYS_REFCURSOR): ignores the return value.
  procedure query2sheet
    ( p_rc             in out sys_refcursor
    , p_column_headers boolean     := true
    , p_sheet          pls_integer := null
    , p_UseXf          boolean     := false
    , p_date_format    varchar2    := 'dd/mm/yyyy'
    , p_title          varchar2    := null
    , p_title_xfid     pls_integer := null
    , p_col            pls_integer := null
    , p_row            pls_integer := null
    , p_autofilter     boolean     := null
    , p_table_style    varchar2    := null
    )
  is
    l_dummy number;
  begin
    l_dummy := query2sheet
                 ( p_rc             => p_rc
                 , p_column_headers => p_column_headers
                 , p_sheet          => p_sheet
                 , p_UseXf          => p_UseXf
                 , p_date_format    => p_date_format
                 , p_title          => p_title
                 , p_title_xfid     => p_title_xfid
                 , p_col            => p_col
                 , p_row            => p_row
                 , p_autofilter     => p_autofilter
                 , p_table_style    => p_table_style
                 );
  end;

  --------------------------------------------------------------------------
  -- Miscellaneous
  --------------------------------------------------------------------------

  -- Toggles XF style resolution on/off. When false, query2sheet uses
  -- pre-resolved style strings for better performance on large data sets.
  procedure setUseXf( p_val boolean := true )
  is
  begin
    g_useXf := p_val;
  end;

  -- Embeds an image (PNG, JPG, GIF, BMP) in a worksheet.
  -- De-duplicates images by CRC hash.  Parses image headers to determine
  -- width and height automatically.
  procedure add_image
    ( p_col pls_integer
    , p_row pls_integer
    , p_img blob
    , p_name varchar2 := ''
    , p_title varchar2 := ''
    , p_description varchar2 := ''
    , p_scale number := null
    , p_sheet pls_integer := null
    , p_width pls_integer := null
    , p_height pls_integer := null
    )
  is
    l_hash raw(4);
    l_image tp_image;
    l_idx pls_integer;
    l_sheet pls_integer := coalesce( p_sheet, workbook.sheets.count );
    l_drawing tp_drawing;
    l_ind number;
    l_len number;
    l_buf raw(32);
    l_hex varchar2(8);
  begin
    l_hash := utl_raw.substr( utl_compress.lz_compress( p_img ), -8, 4 );
    for i in 1 .. workbook.images.count
    loop
      if workbook.images(i).hash = l_hash
      then
        l_idx := i;
        exit;
      end if;
    end loop;
    if l_idx is null
    then
      l_idx := workbook.images.count + 1;
      dbms_lob.createtemporary( l_image.img, true );
      dbms_lob.copy( l_image.img, p_img, dbms_lob.lobmaxsize, 1, 1 );
      l_image.hash := l_hash;
      --
      l_buf := dbms_lob.substr( p_img, 32, 1 );
      if utl_raw.substr( l_buf, 1, 8 ) = hextoraw( '89504E470D0A1A0A' )
      then -- png
        l_ind := 9;
        loop
          l_len := to_number( dbms_lob.substr( p_img, 4, l_ind ), 'xxxxxxxx' );  -- length
          exit when l_len is null or l_ind > dbms_lob.getlength( p_img );
          case rawtohex( dbms_lob.substr( p_img, 4, l_ind + 4 ) ) -- Chunk type
            when '49484452' -- IHDR
            then
              l_image.width  := to_number( dbms_lob.substr( p_img, 4, l_ind + 8 ), 'xxxxxxxx' );
              l_image.height := to_number( dbms_lob.substr( p_img, 4, l_ind + 12 ), 'xxxxxxxx' );
              exit;
            when '49454E44' -- IEND
            then
              exit;
            else
              null;
          end case;
          l_ind := l_ind + 4 + 4 + l_len + 4;  -- Length + Chunk type + Chunk data + CRC
        end loop;
      elsif utl_raw.substr( l_buf, 1, 3 ) = hextoraw( '474946' )
      then -- gif
        l_ind := 14;
        l_buf := utl_raw.substr( l_buf, 11, 1 );
        if utl_raw.bit_and( '80', l_buf ) = '80'
        then
          l_len := to_number( utl_raw.bit_and( '07', l_buf ), 'XX' );
          l_ind := l_ind + 3 * power( 2, l_len + 1 );
        end if;
        loop
          case rawtohex( dbms_lob.substr( p_img, 1, l_ind ) )
            when '21' -- extension
            then
              l_ind := l_ind + 2; -- skip sentinel + label
              loop
                l_len := to_number( dbms_lob.substr( p_img, 1, l_ind ), 'XX' ); -- Block Size
                exit when l_len = 0;
                l_ind := l_ind + 1 + l_len; -- skip Block Size + Data Sub-block
              end loop;
              l_ind := l_ind + 1;           -- skip last Block Size
            when '2C' -- image
            then
              l_buf := dbms_lob.substr( p_img, 4, l_ind + 5 );
              l_image.width  := utl_raw.cast_to_binary_integer( utl_raw.substr( l_buf, 1, 2 ), utl_raw.little_endian );
              l_image.height := utl_raw.cast_to_binary_integer( utl_raw.substr( l_buf, 3, 2 ), utl_raw.little_endian );
              exit;
            else
              exit;
          end case;
        end loop;
      elsif utl_raw.substr( l_buf, 1, 2 ) = hextoraw( 'FFD8' ) -- SOI Start of Image
        and rawtohex( utl_raw.substr( l_buf, 3, 2 ) ) in ( 'FFE0' -- a APP0 jpg
                                                         , 'FFE1' -- a APP1 jpg
                                                         )
      then -- jpg
        l_ind := 5 + to_number( utl_raw.substr( l_buf, 5, 2 ), 'xxxx' );
        loop
          l_buf := dbms_lob.substr( p_img, 4, l_ind );
          l_hex := substr( rawtohex( l_buf ), 1, 4 );
          exit when l_hex in ( 'FFDA' -- SOS Start of Scan
                             , 'FFD9' -- EOI End Of Image
                             )
                 or substr( l_hex, 1, 2 ) != 'FF';
          if l_hex in ( 'FFD0', 'FFD1', 'FFD2', 'FFD3', 'FFD4', 'FFD5', 'FFD6', 'FFD7' -- RSTn
                      , 'FF01'  -- TEM
                      )
          then
            l_ind := l_ind + 2;
          else
            if l_hex = 'FFC0' -- SOF0 (Start Of Frame 0) marker
            then
              l_hex := rawtohex( dbms_lob.substr( p_img, 4, l_ind + 5 ) );
              l_image.width  := to_number( substr( l_hex, 5 ), 'xxxx' );
              l_image.height := to_number( substr( l_hex, 1, 4 ), 'xxxx' );
              exit;
            end if;
            l_ind := l_ind + 2 + to_number( utl_raw.substr( l_buf, 3, 2 ), 'xxxx' );
          end if;
        end loop;
      elsif utl_raw.substr( l_buf, 1, 2 ) = '424D' -- BM
      then -- bmp
        l_image.width  := to_number( utl_raw.reverse( utl_raw.substr( l_buf, 19, 4 ) ), 'XXXXXXXX' );
        l_image.height := to_number( utl_raw.reverse( utl_raw.substr( l_buf, 23, 4 ) ), 'XXXXXXXX' );
      else
        l_image.width  := nvl( p_width, 0 );
        l_image.height := nvl( p_height, 0 );
      end if;
      --
      workbook.images( l_idx ) := l_image;
    end if;
    --
    l_drawing.img_id := l_idx;
    l_drawing.row := p_row;
    l_drawing.col := p_col;
    l_drawing.scale := p_scale;
    l_drawing.name := p_name;
    l_drawing.title := p_title;
    l_drawing.description := p_description;
    workbook.sheets( l_sheet ).drawings( workbook.sheets( l_sheet ).drawings.count + 1 ) := l_drawing;
  end add_image;

  -- Returns the package version string.
  function get_version
  return varchar2
  is
  begin
    return c_version;
  end get_version;
  --
end pck_as_xlsx;
/
