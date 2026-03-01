create or replace package as_xlsx
is
  -------------------------------------------------------------------------------
  -- AS_XLSX - Oracle PL/SQL XLSX Generator (Write-Only)
  -------------------------------------------------------------------------------
  --
  -- Generates Excel .xlsx files (Office Open XML) as BLOB directly from
  -- an Oracle database. Supports multiple sheets, formatting, formulas,
  -- images, comments, hyperlinks, data validation, tables, and optional
  -- password encryption.
  --
  -- Original Author: Anton Scheffer (AMIS)
  -- Original URL:    https://technology.amis.nl/languages/oracle-plsql/create-an-excel-file-with-plsql/
  -- License:         MIT (see LICENSE file)
  --
  -- This is a refactored fork: read functionality, UTL_FILE, and file I/O
  -- have been removed. The finish() function returns a BLOB exclusively.
  --
  -- Changelog (recent):
  --   2025-08-14  Added rotation to alignment
  --   2025-06-22  Re-added formula support
  --   2025-02-01  Fixed BUG with procedure cell (varchar2 value)
  --   2025-01-27  Fixed issue with get_XfId
  --   2024-12-08  Added encryption, tables, fixed multi-sheet combos
  --   2024-11-21  Fixed dates before March 1900, additions to query2sheet
  --   2024-10-18  Fixed add_image for images > 2000 bytes
  --   2026-03-01  Refactored to write-only, split finish() into sub-procedures
  --
  -- Copyright (C) 2011, 2025 by Anton Scheffer
  --
  -- Permission is hereby granted, free of charge, to any person obtaining a
  -- copy of this software and associated documentation files (the "Software"),
  -- to deal in the Software without restriction, including without limitation
  -- the rights to use, copy, modify, merge, publish, distribute, sublicense,
  -- and/or sell copies of the Software, and to permit persons to whom the
  -- Software is furnished to do so, subject to the following conditions:
  --
  -- The above copyright notice and this permission notice shall be included
  -- in all copies or substantial portions of the Software.
  --
  -- THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
  -- OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  -- FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
  -- THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  -- LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
  -- FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
  -- DEALINGS IN THE SOFTWARE.
  -------------------------------------------------------------------------------

  -------------------------------------------------------------------------------
  -- COMPILE-TIME CONFIGURATION
  -------------------------------------------------------------------------------

  -- Set to TRUE to enable password-protected XLSX output via DBMS_CRYPTO.
  -- Requires: GRANT EXECUTE ON dbms_crypto TO <schema>;
  use_dbms_crypto constant boolean := false;

  -------------------------------------------------------------------------------
  -- PUBLIC TYPES
  -------------------------------------------------------------------------------

  -- Cell alignment record used by get_alignment() and formatting procedures.
  type tp_alignment is record
    ( vertical   varchar2(11)   -- 'bottom','center','distributed','justify','top'
    , horizontal varchar2(16)   -- 'center','centerContinuous','distributed','fill',
                                -- 'general','justify','left','right'
    , wrapText   boolean        -- TRUE to wrap text in the cell
    , rotation   number         -- Text rotation in degrees (1..180)
    );

  -------------------------------------------------------------------------------
  -- WORKBOOK LIFECYCLE
  -------------------------------------------------------------------------------

  /**
   * Resets all workbook state (sheets, strings, styles, images).
   * Call before building a new workbook, or rely on finish() which calls
   * clear_workbook automatically after generating the BLOB.
   */
  procedure clear_workbook;

  /**
   * Adds a new worksheet to the workbook.
   *
   * @param p_sheetname      Sheet tab name (max ~31 chars; invalid chars stripped)
   * @param p_tabcolor       Tab color as hex ARGB, e.g. 'FF0000FF' for blue
   * @param p_show_gridlines FALSE to hide gridlines (default: show)
   * @param p_grid_color_idx Index in the default color palette (0..55)
   * @param p_show_headers   FALSE to hide row/column headers
   */
  procedure new_sheet
    ( p_sheetname      varchar2    := null
    , p_tabcolor       varchar2    := null
    , p_show_gridlines boolean     := null
    , p_grid_color_idx pls_integer := null
    , p_show_headers   boolean     := null
    );

  /**
   * Generates the XLSX file and returns it as a BLOB.
   * Automatically calls clear_workbook after generation.
   *
   * @param p_password  Optional password for AES-256 encryption
   *                    (requires use_dbms_crypto = TRUE)
   * @return            The complete XLSX file as BLOB
   *
   * Example:
   *   as_xlsx.new_sheet('Demo');
   *   as_xlsx.cell(1, 1, 'Hello World');
   *   l_blob := as_xlsx.finish;
   */
  function finish
    ( p_password varchar2 := null
    )
  return blob;

  -------------------------------------------------------------------------------
  -- FORMAT HELPERS
  --
  -- These functions register formatting objects in the workbook and return an
  -- ID (pls_integer) that can be passed to cell(), set_row(), set_column(), etc.
  -------------------------------------------------------------------------------

  /**
   * Converts an Oracle date format string to an Excel-compatible format string.
   *
   * @param p_format  Oracle format, e.g. 'dd/mon/yyyy hh24:mi'
   * @return          Excel format, e.g. 'dd/mmm/yyyy hh:mm'
   */
  function OraFmt2Excel
    ( p_format varchar2 := null
    )
  return varchar2;

  /**
   * Registers a custom number/date format and returns its ID.
   *
   * @param p_format  Excel format code, e.g. '#,##0.00', 'dd/mm/yyyy'
   * @return          numFmtId for use in cell() or get_xfid()
   */
  function get_numFmt
    ( p_format varchar2 := null
    )
  return pls_integer;

  /**
   * Registers a font and returns its ID.
   *
   * @param p_name       Font name, e.g. 'Calibri', 'Arial'
   * @param p_family     Font family (default 2 = Swiss)
   * @param p_fontsize   Size in points (default 11)
   * @param p_theme      Theme index (default 1)
   * @param p_underline  TRUE for underlined text
   * @param p_italic     TRUE for italic text
   * @param p_bold       TRUE for bold text
   * @param p_rgb        Font color as hex ARGB, e.g. 'FFFF0000' for red
   * @return             fontId for use in cell() or get_xfid()
   */
  function get_font
    ( p_name      varchar2
    , p_family    pls_integer := 2
    , p_fontsize  number      := 11
    , p_theme     pls_integer := 1
    , p_underline boolean     := false
    , p_italic    boolean     := false
    , p_bold      boolean     := false
    , p_rgb       varchar2    := null
    )
  return pls_integer;

  /**
   * Registers a fill pattern and returns its ID.
   *
   * @param p_patternType  One of: 'none','solid','darkGray','mediumGray',
   *                       'lightGray','gray125','gray0625','darkHorizontal',
   *                       'darkVertical','darkDown','darkUp','darkGrid',
   *                       'darkTrellis','lightHorizontal','lightVertical',
   *                       'lightDown','lightUp','lightGrid','lightTrellis'
   * @param p_fgRGB        Foreground color as hex ARGB
   * @return               fillId for use in cell() or get_xfid()
   */
  function get_fill
    ( p_patternType varchar2
    , p_fgRGB       varchar2 := null
    )
  return pls_integer;

  /**
   * Registers a border style and returns its ID.
   *
   * @param p_top     Top border style (default 'thin')
   * @param p_bottom  Bottom border style (default 'thin')
   * @param p_left    Left border style (default 'thin')
   * @param p_right   Right border style (default 'thin')
   * @param p_rgb     Border color as hex ARGB
   * @return          borderId for use in cell() or get_xfid()
   *
   * Valid border styles: 'none','thin','medium','dashed','dotted','thick',
   *   'double','hair','mediumDashed','dashDot','mediumDashDot',
   *   'dashDotDot','mediumDashDotDot','slantDashDot'
   */
  function get_border
    ( p_top    varchar2 := 'thin'
    , p_bottom varchar2 := 'thin'
    , p_left   varchar2 := 'thin'
    , p_right  varchar2 := 'thin'
    , p_rgb    varchar2 := null
    )
  return pls_integer;

  /**
   * Creates an alignment record.
   *
   * @param p_vertical    'bottom','center','distributed','justify','top'
   * @param p_horizontal  'center','centerContinuous','distributed','fill',
   *                      'general','justify','left','right'
   * @param p_wrapText    TRUE to enable word wrap
   * @param p_rotation    Text rotation in degrees (1..180)
   * @return              tp_alignment record
   */
  function get_alignment
    ( p_vertical   varchar2 := null
    , p_horizontal varchar2 := null
    , p_wrapText   boolean  := null
    , p_rotation   number   := null
    )
  return tp_alignment;

  /**
   * Registers a combined cell format (XF) and returns its ID.
   * Use this to pre-compute a style for repeated application via p_xfId.
   *
   * @param p_numFmtId   Number format ID from get_numFmt()
   * @param p_fontId      Font ID from get_font()
   * @param p_fillId      Fill ID from get_fill()
   * @param p_borderId    Border ID from get_border()
   * @param p_alignment   Alignment from get_alignment()
   * @return              xfId for use in cell(..., p_xfId => ...)
   */
  function get_xfid
    ( p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    )
  return pls_integer;

  -------------------------------------------------------------------------------
  -- CELL DATA
  -------------------------------------------------------------------------------

  /**
   * Writes a numeric value to a cell.
   *
   * @param p_col        Column number (1-based)
   * @param p_row        Row number (1-based)
   * @param p_value      Numeric value
   * @param p_numFmtId   Optional number format ID
   * @param p_fontId     Optional font ID
   * @param p_fillId     Optional fill ID
   * @param p_borderId   Optional border ID
   * @param p_alignment  Optional alignment
   * @param p_sheet      Target sheet (default: last added sheet)
   * @param p_xfId       Pre-computed style ID (overrides individual format params)
   */
  procedure cell
    ( p_col       pls_integer
    , p_row       pls_integer
    , p_value     number
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    );

  /** Writes a text value to a cell. See cell(number) for parameter docs. */
  procedure cell
    ( p_col       pls_integer
    , p_row       pls_integer
    , p_value     varchar2
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    );

  /** Writes a date value to a cell. See cell(number) for parameter docs. */
  procedure cell
    ( p_col       pls_integer
    , p_row       pls_integer
    , p_value     date
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_xfId      pls_integer  := null
    );

  -------------------------------------------------------------------------------
  -- FORMULAS
  -------------------------------------------------------------------------------

  /**
   * Writes a formula that returns a numeric result.
   *
   * @param p_col            Column (1-based)
   * @param p_row            Row (1-based)
   * @param p_formula        Excel formula, e.g. 'SUM(A1:A10)'
   * @param p_default_value  Cached value shown before recalculation
   */
  procedure num_formula
    ( p_col           pls_integer
    , p_row           pls_integer
    , p_formula       varchar2
    , p_default_value number       := null
    , p_numFmtId      pls_integer  := null
    , p_fontId        pls_integer  := null
    , p_fillId        pls_integer  := null
    , p_borderId      pls_integer  := null
    , p_alignment     tp_alignment := null
    , p_sheet         pls_integer  := null
    , p_xfId          pls_integer  := null
    );

  /** Writes a formula that returns a text result. */
  procedure str_formula
    ( p_col           pls_integer
    , p_row           pls_integer
    , p_formula       varchar2
    , p_default_value varchar2     := null
    , p_numFmtId      pls_integer  := null
    , p_fontId        pls_integer  := null
    , p_fillId        pls_integer  := null
    , p_borderId      pls_integer  := null
    , p_alignment     tp_alignment := null
    , p_sheet         pls_integer  := null
    , p_xfId          pls_integer  := null
    );

  /** Writes a formula that returns a date result. */
  procedure date_formula
    ( p_col           pls_integer
    , p_row           pls_integer
    , p_formula       varchar2
    , p_default_value date         := null
    , p_numFmtId      pls_integer  := null
    , p_fontId        pls_integer  := null
    , p_fillId        pls_integer  := null
    , p_borderId      pls_integer  := null
    , p_alignment     tp_alignment := null
    , p_sheet         pls_integer  := null
    , p_xfId          pls_integer  := null
    );

  -------------------------------------------------------------------------------
  -- CELL FEATURES
  -------------------------------------------------------------------------------

  /**
   * Adds a hyperlink to a cell.
   *
   * @param p_col       Column (1-based)
   * @param p_row       Row (1-based)
   * @param p_url       External URL (e.g. 'https://example.com')
   * @param p_value     Display text (defaults to URL if omitted)
   * @param p_sheet     Target sheet
   * @param p_location  Internal workbook location (e.g. 'Sheet2!A1')
   * @param p_tooltip   Hover tooltip text
   */
  procedure hyperlink
    ( p_col      pls_integer
    , p_row      pls_integer
    , p_url      varchar2     := null
    , p_value    varchar2     := null
    , p_sheet    pls_integer  := null
    , p_location varchar2     := null
    , p_tooltip  varchar2     := null
    );

  /**
   * Adds a comment (note) to a cell.
   *
   * @param p_col     Column (1-based)
   * @param p_row     Row (1-based)
   * @param p_text    Comment text
   * @param p_author  Author name (shown in bold above text)
   * @param p_width   Comment box width in pixels (default 150)
   * @param p_height  Comment box height in pixels (default 100)
   * @param p_sheet   Target sheet
   */
  procedure comment
    ( p_col    pls_integer
    , p_row    pls_integer
    , p_text   varchar2
    , p_author varchar2     := null
    , p_width  pls_integer  := 150
    , p_height pls_integer  := 100
    , p_sheet  pls_integer  := null
    );

  -------------------------------------------------------------------------------
  -- SHEET LAYOUT
  -------------------------------------------------------------------------------

  /**
   * Merges a rectangular range of cells.
   *
   * @param p_tl_col  Top-left column
   * @param p_tl_row  Top-left row
   * @param p_br_col  Bottom-right column
   * @param p_br_row  Bottom-right row
   * @param p_sheet   Target sheet
   */
  procedure mergecells
    ( p_tl_col pls_integer
    , p_tl_row pls_integer
    , p_br_col pls_integer
    , p_br_row pls_integer
    , p_sheet  pls_integer := null
    );

  /**
   * Adds a dropdown list validation referencing a cell range.
   *
   * @param p_sqref_col  Column of the cell to validate
   * @param p_sqref_row  Row of the cell to validate
   * @param p_tl_col     Top-left column of the source range
   * @param p_tl_row     Top-left row of the source range
   * @param p_br_col     Bottom-right column of the source range
   * @param p_br_row     Bottom-right row of the source range
   * @param p_style      Error style: 'stop', 'warning', or 'information'
   * @param p_show_error TRUE to show error dialog on invalid input
   */
  procedure list_validation
    ( p_sqref_col   pls_integer
    , p_sqref_row   pls_integer
    , p_tl_col      pls_integer
    , p_tl_row      pls_integer
    , p_br_col      pls_integer
    , p_br_row      pls_integer
    , p_style       varchar2    := 'stop'
    , p_title       varchar2    := null
    , p_prompt      varchar     := null
    , p_show_error  boolean     := false
    , p_error_title varchar2    := null
    , p_error_txt   varchar2    := null
    , p_sheet       pls_integer := null
    );

  /** Adds a dropdown list validation referencing a defined name. */
  procedure list_validation
    ( p_sqref_col    pls_integer
    , p_sqref_row    pls_integer
    , p_defined_name varchar2
    , p_style        varchar2    := 'stop'
    , p_title        varchar2    := null
    , p_prompt       varchar     := null
    , p_show_error   boolean     := false
    , p_error_title  varchar2    := null
    , p_error_txt    varchar2    := null
    , p_sheet        pls_integer := null
    );

  /**
   * Creates a named range (defined name).
   *
   * @param p_tl_col      Top-left column
   * @param p_tl_row      Top-left row
   * @param p_br_col      Bottom-right column
   * @param p_br_row      Bottom-right row
   * @param p_name        Name for the range
   * @param p_sheet       Sheet containing the range
   * @param p_localsheet  Scope (sheet index, 0-based) or NULL for global
   */
  procedure defined_name
    ( p_tl_col     pls_integer
    , p_tl_row     pls_integer
    , p_br_col     pls_integer
    , p_br_row     pls_integer
    , p_name       varchar2
    , p_sheet      pls_integer := null
    , p_localsheet pls_integer := null
    );

  -------------------------------------------------------------------------------
  -- COLUMN & ROW FORMATTING
  -------------------------------------------------------------------------------

  /**
   * Sets the width of a column.
   *
   * @param p_col    Column number (1-based)
   * @param p_width  Width in character units (e.g. 15 ~ 15 characters)
   * @param p_sheet  Target sheet
   */
  procedure set_column_width
    ( p_col   pls_integer
    , p_width number
    , p_sheet pls_integer := null
    );

  /**
   * Sets default formatting for an entire column.
   * Cells in this column inherit these styles unless overridden.
   */
  procedure set_column
    ( p_col       pls_integer
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    );

  /**
   * Sets formatting and/or height for an entire row.
   *
   * @param p_height  Row height in points (e.g. 20.0)
   */
  procedure set_row
    ( p_row       pls_integer
    , p_numFmtId  pls_integer  := null
    , p_fontId    pls_integer  := null
    , p_fillId    pls_integer  := null
    , p_borderId  pls_integer  := null
    , p_alignment tp_alignment := null
    , p_sheet     pls_integer  := null
    , p_height    number       := null
    );

  -------------------------------------------------------------------------------
  -- FREEZE PANES
  -------------------------------------------------------------------------------

  /** Freezes the top N rows (header freeze). */
  procedure freeze_rows
    ( p_nr_rows pls_integer := 1
    , p_sheet   pls_integer := null
    );

  /** Freezes the left N columns. */
  procedure freeze_cols
    ( p_nr_cols pls_integer := 1
    , p_sheet   pls_integer := null
    );

  /** Freezes both rows and columns at the given position. */
  procedure freeze_pane
    ( p_col   pls_integer
    , p_row   pls_integer
    , p_sheet pls_integer := null
    );

  -------------------------------------------------------------------------------
  -- AUTOFILTER & TABLES
  -------------------------------------------------------------------------------

  /**
   * Enables autofilter dropdown buttons on a range.
   * Only one autofilter per sheet is supported.
   */
  procedure set_autofilter
    ( p_column_start pls_integer := null
    , p_column_end   pls_integer := null
    , p_row_start    pls_integer := null
    , p_row_end      pls_integer := null
    , p_sheet        pls_integer := null
    );

  /**
   * Formats a range as an Excel table with banded rows.
   *
   * @param p_style  Table style name, e.g. 'TableStyleMedium2'.
   *                 Valid: TableStyleLight1..21, TableStyleMedium1..28,
   *                        TableStyleDark1..11
   * @param p_name   Table name (auto-generated if NULL)
   */
  procedure set_table
    ( p_column_start pls_integer
    , p_column_end   pls_integer
    , p_row_start    pls_integer
    , p_row_end      pls_integer
    , p_style        varchar2
    , p_name         varchar2    := null
    , p_sheet        pls_integer := null
    );

  /** Sets the sheet tab color (can also be set via new_sheet). */
  procedure set_tabcolor
    ( p_tabcolor varchar2
    , p_sheet    pls_integer := null
    );

  -------------------------------------------------------------------------------
  -- QUERY TO SHEET
  --
  -- Populates a sheet directly from a SQL query or SYS_REFCURSOR.
  -- Supports column headers, titles, autofilters, and table formatting.
  -------------------------------------------------------------------------------

  /** Procedure overload (discards row count). */
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
    );

  /** Procedure overload with SYS_REFCURSOR (discards row count). */
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
    );

  /**
   * Populates a sheet from a SQL string and returns the number of data rows.
   *
   * @param p_sql             SQL query string
   * @param p_column_headers  TRUE to write column headers (default TRUE)
   * @param p_sheet           Target sheet (auto-creates if NULL)
   * @param p_UseXf           TRUE to apply column/row default formatting
   * @param p_date_format     Excel format for date columns (default 'dd/mm/yyyy')
   * @param p_title           Optional title row above the data
   * @param p_title_xfid      Style for the title row
   * @param p_col             Starting column (default 1)
   * @param p_row             Starting row (default 1)
   * @param p_autofilter      TRUE to add autofilter to the header row
   * @param p_table_style     Table style name (adds Excel table formatting)
   * @return                  Number of data rows written
   */
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
  return number;

  /** Function overload accepting a SYS_REFCURSOR. Returns row count. */
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
  return number;

  -------------------------------------------------------------------------------
  -- MISCELLANEOUS
  -------------------------------------------------------------------------------

  /**
   * Toggles whether cell() applies column/row default XF formatting.
   * Used internally by query2sheet; can be called manually if needed.
   */
  procedure setUseXf
    ( p_val boolean := true
    );

  /**
   * Embeds an image in a sheet.
   * Supported formats: PNG, JPG, GIF, BMP. For other formats, provide
   * p_width and p_height manually.
   *
   * @param p_col          Column where image is anchored (1-based)
   * @param p_row          Row where image is anchored (1-based)
   * @param p_img          Image data as BLOB
   * @param p_name         Image name
   * @param p_title        Image title
   * @param p_description  Alt text / description
   * @param p_scale        Scale factor (e.g. 0.5 = half size)
   * @param p_sheet        Target sheet
   * @param p_width        Override width in pixels (for unknown formats)
   * @param p_height       Override height in pixels (for unknown formats)
   */
  procedure add_image
    ( p_col         pls_integer
    , p_row         pls_integer
    , p_img         blob
    , p_name        varchar2    := ''
    , p_title       varchar2    := ''
    , p_description varchar2    := ''
    , p_scale       number      := null
    , p_sheet       pls_integer := null
    , p_width       pls_integer := null
    , p_height      pls_integer := null
    );

  /** Returns the package version string. */
  function get_version
  return varchar2;

end as_xlsx;
/
