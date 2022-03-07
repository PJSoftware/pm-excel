package Excel;

use strict;

our $VERSION = '0.006';
$VERSION = eval $VERSION;   # Correctly convert underscored version numbers

use Win32::OLE;
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;

use Speak qw{ abort whine debug };
use Cwd qw{ cwd };

use File::Spec;
use File::Copy;
use Scalar::Util;

{

my @ValidArgs = qw{ filename existing template password open_pwd readonly autosave recreate };
my %Arg = ();
my %ObjData = ();
my %ClsData = ();

# pseudo-constructor 'new' implies creation of a new excel file
# currently allows (but complains about) "open()" behaviour; will die in future
sub new {
    my $class = shift;
    my $arg = _process_arguments(@_);
    if ($arg->{existing}) {
        whine "If opening an existing file, use open() rather than new()";
        }
    else {
        whine "'readonly' only makes sense for open()" if $arg->{readonly};
        $arg->{recreate} = 1;
        $arg->{autosave} = 1;
        $arg->{readonly} = 0;
        }
    return $class->_create($arg);
    }

# pseudo-constructor 'open' implies we are opening an existing file
sub open {
    my $class = shift;
    my $arg = _process_arguments(@_);
    $arg->{existing} = 1;
    $arg->{template} = '';
    return $class->_create($arg);
    }

sub filename {
    my $self = shift;
    (my $rv = $ObjData{$self}{Filename}) =~ s{\\}{/}g;
    return $rv;
    }

##############################################################################

sub save {
    my $self = shift;
    if ($Arg{$self}{readonly}) {
        whine 'save() request ignored because file set to readonly';
        return;
        }

    if ($ObjData{$self}{Exists}) {
        $ObjData{$self}{Workbook}->Save();
        }
    else {
        $ClsData{App}->{DisplayAlerts} = 0;
        if ($ObjData{$self}{Filename} =~ /\.(xls[bmx]?)$/i) {
            my $ext = $1;
            my %ftype = (
                xls  => xlExcel8,
                xlsb => xlExcel12,
                xlsm => xlOpenXMLWorkbookMacroEnabled,
                xlsx => xlOpenXMLWorkbook,
                );
            $ObjData{$self}{Workbook}->SaveAs($ObjData{$self}{Filename},$ftype{$ext}) ||
                abort "Error while saving '$ObjData{$self}{Filename}': ".Win32::OLE->LastError();
            }
        else {
            abort "Unrecognised file extension";
            }
        $ObjData{$self}{Exists} = 1;
        }
    }

sub close {
    my $self = shift;
    return unless $ObjData{$self}{Workbook};

    $self->save() if $Arg{$self}{autosave};

    $ObjData{$self}{Workbook}->Close(0);  # Do Not Save Changes
    $ObjData{$self}{Workbook} = 0;
    $ClsData{UserCount}--;
    unless ($ClsData{UserCount}) {
        $ClsData{App}->Quit();
        delete $ClsData{App};
        }
    }

sub get_workbook {
    my $self = shift;
    return $ObjData{$self}{Workbook};
    }

sub get_excel {
    my $self = shift;
    return $ClsData{App};
    }

sub sheet_exists {
    my $self = shift;
    my ($sheet_id) = @_;

    my $rv;
    eval { $rv = $self->sheet_by_name($sheet_id); };
    if ($@) {
        eval{ $rv = $self->sheet_by_index($sheet_id); };
        }

    return $rv;
    }

sub sheet_by_name {
    my $self = shift;
    my ($sheet_name) = @_;

    my $wb = $self->get_workbook();
    my $max_sheets = $wb->Sheets()->Count();

    foreach my $idx (1 .. $max_sheets) {
        if ($wb->Sheets($idx)->{Name} eq $sheet_name) {
            debug "Found sheet name '$sheet_name' at index $idx";
            return $wb->Sheets()->Item($idx);
            }
        }
    abort "Invalid sheet name '$sheet_name'";
    }

sub sheet_by_index {
    my $self = shift;
    my ($idx) = @_;

    my $wb = $self->get_workbook();
    my $max_sheets = $wb->Sheets()->Count();

    if (($idx =~ /^\d+$/) && ($idx <= $max_sheets) && ($idx > 0)) {
        debug "Returning sheet by index $idx";
        return $wb->Sheets()->Item($idx);
        }
    else {
        abort "Invalid index '$idx'";
        }
    }

sub DESTROY {
    my $self = shift;

    $self->close();

    foreach my $key (sort keys %{$ObjData{$self}}) {
        debug "Deleting $self data: $key";
        delete $ObjData{$self}{$key};
        }

    delete $Arg{$self};
    }

##### Class-level modules #####

sub objtype {
    my $class = shift;
    my ($obj) = @_;
    my $ref = ref($obj);
    return undef unless $ref;

    return $ref if $obj->isa('Excel');
    return Win32::OLE->QueryObjectType($obj) if $obj->isa('Win32::OLE');
    return 0;
    }

##### Internal use only #####

# _create() is the actual Constructor for this module, called by new() and open()
sub _create {
    my $class = shift;
    my $self = bless \(my $dummy), $class;
    debug "\$self = $self";

    my ($arg_ref) = @_;
    $self->_validate_args($arg_ref);

    $ObjData{$self}{Filename} = $self->_verify_files();
    $ObjData{$self}{Exists} = -f $ObjData{$self}{Filename};
    $ClsData{App} = _connect_to_excel() unless $ClsData{App};
    $ClsData{UserCount}++;

    foreach my $key (sort keys %{$ObjData{$self}}) {        
        debug "ObjData{$key} = '$ObjData{$self}{$key}'";
        }
    foreach my $key (sort keys %{$Arg{$self}}) {        
        debug "Arg{$key} = '$Arg{$self}{$key}'";
        }
    foreach my $key (sort keys %ClsData) {        
        debug "ClsData{$key} = '$ClsData{$key}'";
        }

    if ($ObjData{$self}{Exists}) {
        $ObjData{$self}{Workbook} = $ClsData{App}->Workbooks->Open(
            $ObjData{$self}{Filename}, 0, $Arg{$self}{readonly}, 1, $Arg{$self}{open_pwd}, $Arg{$self}{password}) ||
                abort 'Could not open input file: '.Win32::OLE->LastError();
        }
    else {
        $ObjData{$self}{Workbook} = $ClsData{App}->Workbooks->Add();
        $self->save();
        }

    return $self;
    }

sub _process_arguments {
    my ($arg_ref) = @_;
    my %rv = ();

    if ($arg_ref) {
        if (ref($arg_ref) eq 'HASH') {
            %rv = %{$arg_ref};
            }
        elsif (ref($arg_ref) eq '') {
            $rv{filename} = $arg_ref;
            }
        else {
            abort "Unexpected non-hash reference passed to constructor";
            }
        }
    else {
        abort "Filename must be specified";
        }
    return \%rv;
    }

sub _validate_args {
    my $self = shift;
    my ($arg_ref) = @_;
    %{$Arg{$self}} = ();

    my %is_valid = map { $_ => 1 } @ValidArgs;
    foreach my $arg (keys %{$arg_ref}) {
        abort "Invalid argument '$arg' passed to constructor" unless $is_valid{$arg};
        $Arg{$self}{$arg} = $arg_ref->{$arg};
        }
    foreach my $arg (@ValidArgs) {
        $Arg{$self}{$arg} //= '';
        }

    if (!$Arg{$self}{filename}) { abort "Must specify a filename"; }
    if ($Arg{$self}{template} && !-f $Arg{$self}{template}) { abort "Specified template '$Arg{$self}{template}' not found"; }
    if ($Arg{$self}{readonly}) {
        abort "'autosave' and 'readonly' contradict each other" if $Arg{$self}{autosave};
        abort "'recreate' and 'readonly' contradict each other" if $Arg{$self}{recreate};
        abort "'template' and 'readonly' contradict each other" if $Arg{$self}{template};
        }
    if ($Arg{$self}{existing}) {
        abort "'existing' and 'template' contradict each other" if $Arg{$self}{template};
        }

    foreach my $arg (qw{ existing readonly autosave recreate }) {
        $Arg{$self}{$arg} = $Arg{$self}{$arg} ? 1 : 0;
        }
    }

# Connect to the OLE Excel application.
sub _connect_to_excel {
    my $excel = Win32::OLE->new('Excel.Application', 'quit');
    abort 'Could not start Excel: ' . Win32::OLE->LastError() unless $excel;
    $excel->{Visible} = 0;
    return $excel;
    }

# Ensure that the specified files exist (or can be created!) and are valid filetypes.
sub _verify_files {
    my $self = shift;
    my ($filename,$template);

    $template = _confirm_file($Arg{$self}{template},'template',1) if $Arg{$self}{template};
    $filename = _confirm_file($Arg{$self}{filename},'workbook',$Arg{$self}{existing} && !$Arg{$self}{template});

    if ($Arg{$self}{recreate} || $template) {
        unlink($filename) if -f $filename;
        abort "Unable to delete '$filename'; is somebody in it?" if -f $filename;
        }

    if ($template) {
        copy($template,$filename);
        }

    return $filename;
    }

sub _confirm_file {
    my ($filename,$title,$must_exist) = @_;

    # Valid extension?
    abort "$title '$filename' does not have a recognisable extension"
        unless $filename =~ /.+\.xls[xmb]?$/i;

    # We can fix all other path problems in one step
    $filename = File::Spec->rel2abs($filename);

    if ($must_exist) {
        abort "Required $title '$filename' does not exist"
            unless -f $filename;
        }

    return $filename;
    }

}

1;
