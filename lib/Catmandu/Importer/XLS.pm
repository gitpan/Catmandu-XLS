package Catmandu::Importer::XLS;

use namespace::clean;
use Catmandu::Sane;
use Spreadsheet::ParseExcel;
use Moo;

with 'Catmandu::Importer';

has xls => (is => 'ro', builder => '_build_xls');
has fields => (
    is     => 'rw',
    coerce => sub {
        my $fields = $_[0];
        if (ref $fields eq 'ARRAY') { return $fields }
        if (ref $fields eq 'HASH')  { return [sort keys %$fields] }
        return [split ',', $fields];
    },
);
has _n => (is => 'rw', default => sub { 0 });
has _row_min => (is => 'rw');
has _row_max => (is => 'rw');
has _col_min => (is => 'rw');
has _col_max => (is => 'rw');

sub _build_xls {
    my ($self) = @_;
    my $parser   = Spreadsheet::ParseExcel->new();
    my $xls = $parser->parse( $self->file ) or die $parser->error();

    # process only first worksheet
    $xls = ( $xls->worksheets() )[0];
    ($self->{_row_min}, $self->{_row_max}) = $xls->row_range();
    ($self->{_col_min}, $self->{_col_max}) = $xls->col_range();
    return $xls;
}

sub generator {
    my ($self) = @_;
    sub {
        while ($self->_n <= $self->_row_max) {
            if (!defined $self->fields) {
                $self->fields([$self->_get_row]);
                $self->{_n}++;
            }
            else{
                my @data = $self->_get_row();
                $self->{_n}++;
                my @fields = @{$self->fields()};
                my %hash = map { 
                    my $key = shift @fields;
                    defined $_  ? ($key => $_) : ()
                    } @data;
                return \%hash;
            }
        }
        return;
    }
}

sub _get_row {
    my ($self) = @_;
    my @row;
    for my $col ( $self->_col_min .. $self->_col_max ) {
        my $cell = $self->xls->get_cell( $self->_n, $col );
        if ($cell) {
            push(@row,$cell->unformatted());
        }
        else{
            push(@row, undef);            
        }
    }
    return @row;
}

=head1 NAME

Catmandu::Importer::XLS - Package that imports XLS files

=head1 SYNOPSIS

    use Catmandu::Importer::XLS;

    my $importer = Catmandu::Importer::XLS->new(file => "./t/test.xls");

    my $n = $importer->each(sub {
        my $hashref = $_[0];
        # ...
    });

=head1 METHODS

=head2 new(file => $filename [, fields => \@fields])

Create a new XLS importer for $filename. Use STDIN when no filename is given. The
object fields are read from the first XLS row or given via the 'fields' parameter.

Only the first worksheet from the Excel workbook is imported.

=head2 count

=head2 each(&callback)

=head2 ...

Every L<Catmandu::Importer> is a L<Catmandu::Iterable> all its methods are
inherited. The Catmandu::Importer::XLS methods are not idempotent: XLS streams
can only be read once.

=head1 SEE ALSO

L<Catmandu::Iterable>

=cut

1;
