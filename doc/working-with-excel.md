# Working With Excel

## Overview

### Initial Setup

```perl
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';
```

### Starting Excel

```perl
my $excel = Win32::OLE->new('Excel.Application','quit') ||
    croak "Could not start Excel: ".Win32::OLE->LastError()."\n";
$excel->{Visible} = 0;
```

### Opening Spreadsheet

```perl
    my $xlFile = "$path_to\\filename.xls";
    my $workbook = $excel->Workbooks->Open($xlFile) ||
        croak "Could not open input file: ".Win32::OLE->LastError()."\n";
```

### To open READ-ONLY

```perl
    my $xlFile = "$path_to/filename.xls";
    my $workbook = $excel->Workbooks->Open($xlFile,0,1) ||
        croak "Could not open input file: ".Win32::OLE->LastError()."\n";
```

### To open in Write mode with specified Password

```perl
    my $xlFile = "$path_to/filename.xls";
    my $workbook = $excel->Workbooks->Open($xlFile,0,0,1,'',$WritePassword) ||
        croak "Could not open input file: ".Win32::OLE->LastError()."\n";
```

### All Open() Parameters

```perl
my $workbook = $excel->Workbooks->Open(
    $Filename,$UpdateLinks,$ReadOnly,$Format,$Password,
    $WriteResPassword,$IgnoreReadOnlyRecommended,$Origin,
    $Delimiter,$Editable,$Notify,$Converter,$AddToMRU,
    $Local,$CorruptLoad);
```

### Accessing Worksheet

```perl
my $worksheet = $workbook->Sheets('SheetName');
```

### Adding Workbook and Worksheets

```perl
my $workbook = $excel->Workbooks->Add();
my $ws1 = $workbook->Sheets(1);
$workbook->Worksheets->Add();
my $ws2 = $workbook->Sheets(2);
$ws1->{Name} = 'Sheet1';
$ws2->{Name} = 'Additional Sheet';
```

### Managing Multiple Worksheets

```perl
my @worksheet = ();
my $wsReq = 10;
while ($workbook->Worksheets->Count < $wsReq) {
    $workbook->Worksheets->Add();
    }
for (my $wsIdx = 1; $wsIdx <= $wsReq; $wsIdx++) {
    $worksheet[$wsIdx-1] = $workbook->Sheets($wsIdx);
    $worksheet[$wsIdx-1]->{Name} = "Sheet $wsIdx";
    }
```

### Workbook Data

```perl
my $numSheets = $workbook->Worksheets->Count();

$workbook->ResetColors();
```

### Worksheet Data

```perl
my $numRows = $worksheet->UsedRange->Rows->Count();
my $numCols = $worksheet->UsedRange->Columns->Count();
```

### Accessing Worksheet Cell 'B5'

```perl
$worksheet->Range("B5")->{Value} = "This is Cell B5";
$worksheet->Cells(5,2)->{Value} = "This is Cell B5";
```

### Reading Worksheet Data

```perl
my $value       = $worksheet->Range("B5")->{Value} || '';
my $formula     = $worksheet->Range("B5")->{Formula} || '';

my $shade       = $worksheet->Range("B5")->Interior->{ColorIndex};
my $pattern     = $worksheet->Range("B5")->Interior->{Pattern};
my $patColor    = $worksheet->Range("B5")->Interior->{PatternColorIndex};

my $boldness    = $worksheet->Range("B5")->Font->{Bold};
my $textSize    = $worksheet->Range("B5")->Font->{Size};
my $hAlign      = $worksheet->Range("B5")->{HorizontalAlignment};
my $nFormat     = $worksheet->Range("B5")->{NumberFormat};
```

### Changing Worksheet Data

```perl
    $worksheet->Cells($row,$col)->Interior->{ColorIndex}        = 34;
    $worksheet->Cells($row,$col)->Interior->{Pattern}           = xlGray75;
    $worksheet->Cells($row,$col)->Interior->{PatternColorIndex} = 2;
```

### Miscellaneous Formatting

```perl
    $worksheet->Range("A1:B1")->Font->{'Bold'} = 1;
    $worksheet->Columns("A:B")->EntireColumn->AutoFit;
    $excel->Selection->AutoFilter;
    $worksheet->Range("A2")->Select;
    $excel->ActiveWindow->{'FreezePanes'} = 1;
```

### Shutting Down

```perl
    $workbook->Save();
    $workbook->Close();
     or $workbook->Close(1);  # Save
     or $workbook->Close(0);  # Do not save
    ## These originally used xlSaveChanges and xlDoNotSaveChanges but this is not correct!
    ## Close() is looking for a boolean to tell it whether to save or not!
    $excel->Quit();
```

From [http://www.tek-tips.com/faqs.cfm?fid=6715](http://www.tek-tips.com/faqs.cfm?fid=6715)

```perl
use OLE;
use Win32::OLE::Const "Microsoft Excel";
```

## DEFINE EXCEL

```perl
$excel = CreateObject OLE "Excel.Application";
```

## MAKE EXCEL VISIBLE

```perl
$excel->{Visible} = 1;
```

## ADD NEW WORKBOOK

```perl
$workbook = $excel->Workbooks->Add;
$sheet = $workbook->Worksheets("Sheet1");
$sheet->Activate;
```

## OPEN EXISTING WORKBOOK

```perl
$workbook = $excel->Workbooks->Open("$file_name");
$sheet    = $workbook->Worksheets(1)->{Name};
$sheet = $workbook->Worksheets($sheet);
$sheet->Activate;
```

## ACTIVATE EXISTING WORKBOOK

```perl
$excel->Windows("Book1")->Activate;
$workbook = $excel->Activewindow;
$sheet = $workbook->Activesheet;
```

## CLOSE WORKBOOK

```perl
$workbook->Close;
```

## ADD NEW WORKSHEET

```perl
$workbook->Worksheets->Add({After => $workbook->Worksheets($workbook->Worksheets->{Count})});
```

## CHANGE WORKSHEET NAME

```perl
$sheet->{Name} = "Name of Worksheet";
```

## PRINT VALUE TO CELL

```perl
$sheet->Range("A1")->{Value} = 1234;
```

## SUM FORMULAS

```perl
$sheet->Range("A3")->{FormulaR1C1} = "=SUM(R[-2]C:R[-1]C)"; # Sum rows
$sheet->Range("C1")->{FormulaR1C1} = "=SUM(RC[-2]:RC[-1])"; # Sum columns
```

## RETRIEVE VALUE FROM CELL

```perl
$data = $sheet->Range("G7")->{Value};
```

## FORMAT TEXT

```perl
$sheet->Range("G7:H7")->Font->{Bold}        = "True";
$sheet->Range("G7:H7")->Font->{Italic}      = "True";
$sheet->Range("G7:H7")->Font->{Underline}   = xlUnderlineStyleSingle;
$sheet->Range("G7:H7")->Font->{Size}        = 8;
$sheet->Range("G7:H7")->Font->{Name}        = "Arial";
$sheet->Range("G7:H7")->Font->{ColorIndex}  = 4;

$sheet->Range("G7:H7")->{NumberFormat} = "\@";                              # Text
$sheet->Range("A1:H7")->{NumberFormat} = "\$#,##0.00";                      # Currency
$sheet->Range("G7:H7")->{NumberFormat} = "\$#,##0.00_);[Red](\$#,##0.00)";  # Currency - red negatives
$sheet->Range("G7:H7")->{NumberFormat} = "0.00*);[Red](0.00)";              # Numbers with decimals
$sheet->Range("G7:H7")->{NumberFormat} = "#,##0";                           # Numbers with commas
$sheet->Range("G7:H7")->{NumberFormat} = "#,##0*);[Red](#,##0)";            # Numbers with commas - red negatives
$sheet->Range("G7:H7")->{NumberFormat} = "0.00%";                           # Percents
$sheet->Range("G7:H7")->{NumberFormat} = "m/d/yyyy";                        # Dates
```

## ALIGN TEXT

```perl
$sheet->Range("G7:H7")->{HorizontalAlignment} = xlHAlignCenter; # Center text;
$sheet->Range("A1:A2")->{Orientation} = 90;                     # Rotate text
```

## SET COLUMN WIDTH/ROW HEIGHT

```perl
$sheet->Range('A:A')->{ColumnWidth} = 9.14;
$sheet->Range("8:8")->{RowHeight} = 30;
$sheet->Range("G:H")->{Columns}->Autofit;
```

## FIND LAST ROW/COLUMN WITH DATA

```perl
$last_row = $sheet->UsedRange->Find({What => "*", SearchDirection => xlPrevious, SearchOrder => xlByRows})->{Row};
$last_col = $sheet->UsedRange->Find({What => "\*", SearchDirection => xlPrevious, SearchOrder => xlByColumns})->{Column};
```

## ADD BORDERS

```perl
$sheet->Range("A3:I3")->Borders(xlEdgeBottom)->{LineStyle}  = xlDouble;
$sheet->Range("A3:I3")->Borders(xlEdgeBottom)->{Weight} = xlThick;
$sheet->Range("A3:I3")->Borders(xlEdgeBottom)->{ColorIndex} = 1;
$sheet->Range("A3:I3")->Borders(xlEdgeLeft)->{LineStyle} = xlContinuous;
$sheet->Range("A3:I3")->Borders(xlEdgeLeft)->{Weight}     = xlThin;
$sheet->Range("A3:I3")->Borders(xlEdgeTop)->{LineStyle} = xlContinuous;
$sheet->Range("A3:I3")->Borders(xlEdgeTop)->{Weight}     = xlThin;
$sheet->Range("A3:I3")->Borders(xlEdgeBottom)->{LineStyle} = xlContinuous;
$sheet->Range("A3:I3")->Borders(xlEdgeBottom)->{Weight}     = xlThin;
$sheet->Range("A3:I3")->Borders(xlEdgeRight)->{LineStyle} = xlContinuous;
$sheet->Range("A3:I3")->Borders(xlEdgeRight)->{Weight}     = xlThin;
$sheet->Range("A3:I3")->Borders(xlInsideVertical)->{LineStyle} = xlContinuous;
$sheet->Range("A3:I3")->Borders(xlInsideVertical)->{Weight}     = xlThin;
$sheet->Range("A3:I3")->Borders(xlInsideHorizontal)->{LineStyle} = xlContinuous;
$sheet->Range("A3:I3")->Borders(xlInsideHorizontal)->{Weight} = xlThin;
```

## PRINT SETUP

```perl
$sheet->PageSetup->{Orientation}  = xlLandscape;
$sheet->PageSetup->{Order} = xlOverThenDown;
$sheet->PageSetup->{LeftMargin}   = .25;
$sheet->PageSetup->{RightMargin} = .25;
$sheet->PageSetup->{BottomMargin} = .5;
$sheet->PageSetup->{CenterFooter} = "Page &P of &N";
$sheet->PageSetup->{RightFooter}  = "Page &P of &N";
$sheet->PageSetup->{LeftFooter} = "Left\nFooter";
$sheet->PageSetup->{Zoom}         = 75;
$sheet->PageSetup->FitToPagesWide = 1;
$sheet->PageSetup->FitToPagesTall = 1;
```

## ADD PAGE BREAK

```perl
$excel->ActiveWindow->SelectedSheets->HPageBreaks->Add({Before => $sheet->Range("3:3")});
```

## HIDE COLUMNS

```perl
$sheet->Range("G:H")->EntireColumn->{Hidden} = "True";
```

## MERGE CELLS

```perl
$sheet->Range("H10:J10")->Merge;
```

## INSERT PICTURE

```perl
$sheet->Pictures->Insert("picture_name");               # Insert in upper-left corner
$excel->ActiveSheet->Pictures->Insert("picture_name");  # Insert in active cell
```

## GROUP ROWS

```perl
$sheet->Range("7:8")->Group;
```

## ACTIVATE CELL

```perl
$sheet->Range("A2")->Activate;
```

## FREEZE PANES

```perl
$excel->ActiveWindow->{FreezePanes} = "True";
```

## DELETE SHEET

```perl
$sheet->Delete;
```

## SAVE AND QUIT

```perl
$excel->{DisplayAlerts} = 0; # This turns off the "This file already exists" message.
$workbook->SaveAs ("C:\\file_name.xls");
$excel->Quit;
```
