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
    $module = 'Excel::Utility'; $testver = '0.001';
    @subs = qw{ col_number_to_letter };
    use_ok($module, @subs, $testver);
    }
like ($module->VERSION, qr/^$testver/, 'Test version matches module version');

#~ use My::Test qw{ test_sandbox };
#~ my $SANDBOX = test_sandbox($module);

test_subroutines();
test_cntl();

done_testing();

##############################################################################

sub test_subroutines {
    foreach my $sub (@subs) {
        can_ok($module, $sub);
        }
    }

sub test_cntl {
    subtest 'col_number_to_letter()' => sub {
        throws_ok { col_number_to_letter() }    qr/Expecting an integer value/, "correctly throws exception on null input";
        throws_ok { col_number_to_letter(1.1) } qr/Expecting an integer value/, "correctly throws exception on non-integer input";
        throws_ok { col_number_to_letter(-1) }  qr/Expecting a positive, non-zero integer/, "correctly throws exception on negative input";
        throws_ok { col_number_to_letter(0) }   qr/Expecting a positive, non-zero integer/, "correctly throws exception on zero input";
        throws_ok { col_number_to_letter('A') } qr/Expecting an integer value/, "correctly throws exception on non-numeric input";

        is(col_number_to_letter(1),     'A',    "correctly converts 1 => A");
        is(col_number_to_letter(26),    'Z',    "correctly converts 26 => Z");
        is(col_number_to_letter(27),    'AA',   "correctly converts 27 => AA");
        is(col_number_to_letter(52),    'AZ',   "correctly converts 52 => AZ");
        is(col_number_to_letter(53),    'BA',   "correctly converts 53 => BA");
        is(col_number_to_letter(392),   'OB',   "correctly converts 392 => OB");
        is(col_number_to_letter(702),   'ZZ',   "correctly converts 702 => ZZ");
        is(col_number_to_letter(703),   'AAA',  "correctly converts 703 => AAA");
        is(col_number_to_letter(16384), 'XFD',  "correctly converts max column");

        throws_ok { col_number_to_letter(16385) } qr/Maximum allowable column exceeded/, "correctly throws exception on column greater than max allowable";
        };
    }
