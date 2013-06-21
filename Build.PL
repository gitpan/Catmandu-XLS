use strict;
use warnings;
use Module::Build;

my $builder = Module::Build->new(
    module_name         => 'Catmandu::XLS',
    license             => 'perl',
    build_requires => {
        'Test::Exception' => 0,
        'Test::More' => 0,
    },
    requires => {
        'perl' => '5.10.1',
        'Catmandu' => '0.2003',
        'namespace::clean' => '0.24',
        'Spreadsheet::WriteExcel' => '2.37',
    },
    add_to_cleanup     => [ 'Catmandu-XLS-*' ],
    create_makefile_pl => 'traditional',
);

$builder->create_build_script;
