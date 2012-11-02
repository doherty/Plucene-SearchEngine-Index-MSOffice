package Plucene::SearchEngine::Index::DOC;
use strict;
use warnings;
# VERSION
# ABSTRACT: a Plucene backend for indexing Microsoft Word documents
use parent qw(Plucene::SearchEngine::Index::Text);

use IPC::Run3;
use File::Temp;

__PACKAGE__->register_handler('application/doc', '.doc');

=head1 DESCRIPTION

This backend analyzes a DOC file for its textual content (using C<antiword>).

=head1 METHODS

=head2 gather_data_from_file

Overrides the method from L<Plucene::SearchEngine::Index::Text>
to provide DOC parsing.

=cut

sub gather_data_from_file {
    my ($self, $filename) = @_;
    return unless $filename =~ m/\.doc$/;

    my $tmp_txt = File::Temp->new();
    run3 ['antiword', $filename],
        \undef,     # stdin is /dev/null
        $tmp_txt,   # some temporary file
        undef;      # inherit the parent's stderr
    $self->gather_data_from_file( $tmp_txt->filename );
    return $self;
}

1;
