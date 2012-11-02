package Plucene::SearchEngine::Index::PPT;
use strict;
use warnings;
# VERSION
# ABSTRACT: a Plucene backend for indexing Microsoft Powerpoint presentations
use parent qw(Plucene::SearchEngine::Index::HTML);

use IPC::Run3;
use File::Temp;

__PACKAGE__->register_handler('text/ppt', '.ppt');

=head1 DESCRIPTION

This backend analysis a PPT file. The module use the tool called
ppthtml, provided by xlhtml packges available from
L<http://chicago.sourceforge.net/xlhtml/>, or your operating
system's package manager.

B<This code is not currently actively maintained.>

=over 3

=item text

The text part of the PPT

=item link

A list of links in the HTML

=back

Additionally, any C<META> tags are turned into Plucene fields.

=head1 METHODS

=head2 gather_data_from_file

Overrides the method from L<Plucene::SearchEngine::Index::HTML>
to provide PPT parsing.
=cut

sub gather_data_from_file {
    my ($self, $filename) = @_;
    return unless $filename =~ m/\.ppt$/;

    my $tmp_html = File::Temp->new();
    run3 ['ppthtml', $filename],
        \undef,     # redirect from /dev/null
        $tmp_html,  # write to a temp file
        undef;      # inherit the parent's stderr

    $self->gather_data_from_file( $tmp_html->filename );
    return $self;
}

1;
