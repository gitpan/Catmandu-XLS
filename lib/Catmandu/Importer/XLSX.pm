package Catmandu::Importer::XLSX;

use namespace::clean;
use Catmandu::Sane;
use Encode qw(decode);
use Spreadsheet::XLSX;
use Moo;

with 'Catmandu::Importer';

has xlsx => (is => 'ro', builder => '_build_xlsx');
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

sub _build_xlsx {
    my ($self) = @_;
    my $xlsx = Spreadsheet::XLSX->new( $self->file ) or die "Could not open file " . $self->file;

    # process only first worksheet
    $xlsx = $xlsx->{Worksheet}->[0];
    $self->{_col_min} = $xlsx->{MinCol};
    $self->{_col_max} = $xlsx->{MaxCol};
    $self->{_row_min} = $xlsx->{MinRow};
    $self->{_row_max} = $xlsx->{MaxRow};
    return $xlsx;
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
        my $cell = $self->xlsx->{Cells}[$self->_n][$col];
        if ($cell) {
            push(@row, decode('UTF-8',$cell->{Val}));
        }
        else{
            push(@row, undef);            
        }
    }
    return @row;
}

=head1 NAME

Catmandu::Importer::XLSX - Package that imports XLSX files

=head1 SYNOPSIS

    use Catmandu::Importer::XLSX;

    my $importer = Catmandu::Importer::XLSX->new(file => "./t/test.xlsx");

    my $n = $importer->each(sub {
        my $hashref = $_[0];
        # ...
    });

=head1 METHODS

=head2 new(file => $filename [, fields => \@fields"])

Create a new XLSX importer for $filename. Use STDIN when no filename is given. The
object fields are read from the XLS header line or given via the 'fields' parameter.

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
