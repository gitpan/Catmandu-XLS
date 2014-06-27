use strict;
use warnings;
use Test::More;

use Catmandu::Exporter::XLS;
use Catmandu::Exporter::XLSX;
use Catmandu::Importer::XLS;
use Catmandu::Importer::XLSX;
use File::Temp qw(tempfile);
use IO::File;
use Encode qw(encode);

my @rows = (
    {'number' => '1', 'letter' => 'a'},
    {'number' => '2', 'letter' => 'b'},
    {'number' => '3', 'letter' => 'c'},
);

# XLS

my ( $fh, $filename ) = tempfile();
my $exporter = Catmandu::Exporter::XLS->new(
    fh => $fh,
    header => 1,
    fields => 'number,letter',
);

isa_ok($exporter, 'Catmandu::Exporter::XLS');

for my $row (@rows) {
    $exporter->add($row);
}

$exporter->commit();

close($fh);

my $importer = Catmandu::Importer::XLS->new( file => $filename );
my $rows = $importer->to_array();

is_deeply ($rows->[0], {number => 1, letter => 'a'}, 'first row');
is_deeply ($rows->[1], {number => 2, letter => 'b'}, 'second row');
is_deeply ($rows->[2], {number => 3, letter => 'c'}, 'third row');

# XLSX

( $fh, $filename ) = tempfile();
$exporter = Catmandu::Exporter::XLSX->new(
    fh => $fh,
    header => 1,
    fields => 'number,letter',
);

isa_ok($exporter, 'Catmandu::Exporter::XLSX');

for my $row (@rows) {
    $exporter->add($row);
}

$exporter->commit();

close($fh);

$importer = Catmandu::Importer::XLSX->new( file => $filename );
$rows = $importer->to_array();

is_deeply ($rows->[0], {number => 1, letter => 'a'}, 'first row');
is_deeply ($rows->[1], {number => 2, letter => 'b'}, 'second row');
is_deeply ($rows->[2], {number => 3, letter => 'c'}, 'third row');


done_testing;