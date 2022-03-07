package Excel::Utility;

use strict;
our $VERSION = '0.001';
$VERSION = eval $VERSION;

use parent qw{ Exporter };
our @EXPORT_OK = qw{
    col_number_to_letter
    };

use Speak qw{ abort };

sub col_number_to_letter {
    my ($col) = @_;
    $col //= '';
    abort "Expecting an integer value" unless $col =~ /^-?\d+$/;
    abort "Expecting a positive, non-zero integer" if $col < 1;
    abort "Maximum allowable column exceeded" if $col > 16384;

    my $next = int(($col-1)/26);
    my $conv = $col - ($next * 26);

    return ($next ? col_number_to_letter($next) : '') . chr(ord('A')+$conv-1);
    }

1;
