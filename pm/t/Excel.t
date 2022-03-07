use Test::More;
use Test::Exception;

# Automatically determine libpath
my $LIB;
BEGIN {
    use Cwd;
    ($LIB = cwd()) =~ s{test_automation(/.+)?}{perl_site_lib};
    unshift @INC, $LIB;
    }

use strict;
use warnings;

my ($module,$testver,@subs);
BEGIN {
    $module = 'Excel'; $testver = eval '0.006';
    use_ok($module,, $testver);
    }
like ($module->VERSION, qr/^$testver/, 'Test version matches module version');

use My::Test qw{ test_sandbox };
my $SANDBOX = test_sandbox($module);
my $nosuchfile = "$SANDBOX/NoFileExists.xls";
my $testfile = "$SANDBOX/ExcelTestFile.xls";

my $obj = test_constructor();
test_workbook($obj);
test_filename($obj);
test_objtype($obj);
test_sheets($obj);

done_testing();

##############################################################################

sub test_constructor {
    my ($rv,$temp);

    throws_ok { $module->open() } qr/Filename must be specified/, "open() dies if no filename provided";
    throws_ok { $module->open($nosuchfile) } qr/Required workbook.*does not exist/, "open() dies if no file exists";
    is( -f $nosuchfile,undef,"open() does NOT autocreate file");

    throws_ok { $module->open({ filename => $nosuchfile, template => $testfile }) }
        qr/Required workbook.*does not exist/, "open() ignores template; looking for specified file";

    throws_ok { $module->new() } qr/Filename must be specified/, "new() dies if no filename provided";
    $temp = $module->new($nosuchfile);
    isa_ok($temp,$module, "new() autocreates worksheet for non-existent file");
    is( -f $nosuchfile,1,"new() autocreates file");
    $temp->close();
    unlink($nosuchfile);

    throws_ok { $module->new({ filename => "$SANDBOX/not_a_file.xls", template => $nosuchfile }) }
        qr/Specified template.*not found/, "new() dies if specified template not found";

    $rv = $module->open($testfile);
    return $rv;
    }

sub test_workbook {
    my ($obj) = @_;

    my $wb = $obj->get_workbook();
    isa_ok( $wb, 'Win32::OLE');
    is(Win32::OLE->QueryObjectType($wb),'_Workbook','get_workbook() returns _Workbook object');
    }

sub test_filename {
    my ($obj) = @_;

    is($obj->filename(),$testfile,'filename() returns correct path to spreadsheet');
    }

sub test_objtype {
    my ($obj) = @_;
    is(Excel->objtype("scalar"), undef, 'Correctly detects scalar variable');
    is(Excel->objtype($obj), 'Excel', 'Correctly detects Excel object');
    is(Excel->objtype($obj->get_workbook()),'_Workbook',    'Correctly detects Workbook object');
    is(Excel->objtype($obj->get_excel()),   '_Application', 'Correctly detects Excel_Application object');
    }

sub test_sheets {
    my ($obj) = @_;

    my $sh1 = $obj->sheet_exists('Sheet1');
    my $sh2 = $obj->sheet_exists('NamedSheet');
    my $sh3 = $obj->sheet_exists(3);

    is( Excel->objtype($sh1), '_Worksheet','correctly detected unnamed "Sheet1" sheet');
    is( Excel->objtype($sh2), '_Worksheet','correctly detected named sheet');
    is( Excel->objtype($sh3), '_Worksheet','correctly returns sheet from numeric index');
    is($sh3->{Name},'2','indexed sheet has confusing name, as expected');

    is( $obj->sheet_exists('Sheet3'),undef,'correctly returns undef without throwing exception');

    # Poorly named sheet (sheet name matches index of different sheet!)
    my $x1 = $obj->sheet_exists(2);
    is($x1->{Name},'2','sheet name takes precedence over sheet index');

    my $x2 = $obj->sheet_by_index(2);
    is($x2->{Name},'NamedSheet','sheet_by_index() works as intended');

    my $x3 = $obj->sheet_by_name(2);
    is($x1->{Name},'2','sheet_by_name() works as intended');

    throws_ok { $obj->sheet_by_name(3) } qr/Invalid sheet name/, 'sheet_by_name() throws exception rather than interpret index';
    throws_ok { $obj->sheet_by_name('NoSuchSheet') } qr/Invalid sheet name/, 'sheet_by_name() throws exception on unknown sheet';

    throws_ok { $obj->sheet_by_index('Sheet1') } qr/Invalid index/, 'sheet_by_name() throws exception on non-numeric index';
    throws_ok { $obj->sheet_by_index(4) } qr/Invalid index/, 'sheet_by_name() throws exception if index outside allowable range';
    }
