use strict;
use warnings;
use utf8;
use Test::More;
use Test::Exception;
use Catmandu::Importer::XLS;
use Catmandu::Importer::XLSX;

# XLS

my $importer = Catmandu::Importer::XLS->new( file => './t/test.xls');

isa_ok($importer, 'Catmandu::Importer::XLS');
can_ok($importer, 'each');

my $rows = $importer->to_array();

is_deeply ($rows->[0], {Column1 => 1,Column2 => 'a',Column3 => 0.01}, 'first row');
is_deeply ($rows->[1], {Column1 => 2,Column2 => 'b',Column3 => 2.5}, 'second row');
is_deeply ($rows->[-1], {Column1 => 27,Column3 => 'Ümlaut'}, 'last row');

# XLSX

$importer = Catmandu::Importer::XLSX->new( file => './t/test.xlsx');

isa_ok($importer, 'Catmandu::Importer::XLSX');
can_ok($importer, 'each');

$rows = $importer->to_array();

is_deeply ($rows->[0], {Column1 => 1,Column2 => 'a',Column3 => 0.01}, 'first row');
is_deeply ($rows->[1], {Column1 => 2,Column2 => 'b',Column3 => 2.5}, 'second row');
is_deeply ($rows->[-1], {Column1 => 27,Column3 => 'Ümlaut'}, 'last row');

done_testing;