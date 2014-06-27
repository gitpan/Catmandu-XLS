package Catmandu::Exporter::XLSX;

use namespace::clean;
use Catmandu::Sane;
use Excel::Writer::XLSX;
use Moo;

with 'Catmandu::Exporter';

has xlsx => (is => 'ro', lazy => 1, builder => '_build_xlsx');
has worksheet => (is => 'ro', lazy => 1, builder => '_build_worksheet');
has header => (is => 'ro', default => sub { 1 });
has fields => (
    is     => 'rw',
    coerce => sub {
        my $fields = $_[0];
        if (ref $fields eq 'ARRAY') { return $fields }
        if (ref $fields eq 'HASH')  { return [sort keys %$fields] }
        return [split ',', $fields];
    },
);

our $VERSION = '0.02';

sub _build_xlsx {
    my $xlsx = Excel::Writer::XLSX->new($_[0]->fh);
    $xlsx;
}

sub _build_worksheet {
    $_[0]->xlsx->add_worksheet;
}

sub encoding { ':raw' }

sub add {
    my ($self, $data) = @_;
    my $header = $self->header;
    my $fields = $self->fields || $self->fields($data);
    my $worksheet = $self->worksheet;
    my $n = $self->count;
    if ($header) {
        if ($n == 0) {
            for (my $i = 0; $i < @$fields; $i++) {
                my $field = $fields->[$i];
                $field = $header->{$field} if ref $header && defined $header->{$field};
                $worksheet->write_string($n, $i, $field);
            }
        }
        $n++;
    }
    for (my $i = 0; $i < @$fields; $i++) {
        $worksheet->write_string($n, $i, $data->{$fields->[$i]} // "");
    }
}

sub commit {
    $_[0]->xlsx->close;
}

=head1 NAME

Catmandu::Exporter::XLSX - a XLSX exporter

=head1 VERSION

Version 0.01

=head1 SYNOPSIS

    use Catmandu::Exporter::XLSX;

    my $exporter = Catmandu::Exporter::XLSX->new(
				file => 'output.xlsx',
				fix => 'myfix.txt'
				header => 1);

    $exporter->fields("f1,f2,f3");

    # add custom header labels
    $exporter->header({f2 => 'field two'});

    $exporter->add_many($arrayref);
    $exporter->add_many($iterator);
    $exporter->add_many(sub { });

    $exporter->add($hashref);

    $exporter->commit;

    printf "exported %d objects\n" , $exporter->count;

=head1 METHODS

=head2 new(header => 0|1|HASH, fields => ARRAY|HASH|STRING)

Creates a new Catmandu::Exporter::XLSX. A header line with field names will be
included if C<header> is set. Field names can be read from the first item
exported or set by the fields argument (see: C<fields>).

=head2 fields($arrayref)

Set the field names by an ARRAY reference.

=head2 fields($hashref)

Set the field names by the keys of a HASH reference.

=head2 fields($string)

Set the fields by a comma delimited string.

=head2 header(1)

Include a header line with the field names

=head2 header($hashref)

Include a header line with custom field names

=head2 commit

Commit the changes and close the XLSX.

=head1 SEE ALSO

L<Catmandu::Exporter::XLS>, L<Catmandu::Importer::XLS>, <Catmandu::Importer::XLSX>.

=cut

1;
