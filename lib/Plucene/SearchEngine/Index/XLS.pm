package Plucene::SearchEngine::Index::XLS;
use strict;
use warnings;
# VERSION
# ABSTRACT: a Plucene backend for indexing Microsoft Excel spreadsheets

use parent qw(Plucene::SearchEngine::Index::Base);

__PACKAGE__->register_handler('application/xls', '.xls');
use File::Temp qw/tmpnam/;
use Spreadsheet::ParseExcel;


=head1 NAME

Plucene::SearchEngine::Index::Xls - Backend for plain text files

=head1 DESCRIPTION

This backend converts the .xls file into text file and the text file
is used similar to Text.pm module.


=head1 METHODS

=head2 gather_data_from_file

Overrides the method from L<Plucene::SearchEngine::Index::Base>
to provide XLS parsing.

=cut

sub gather_data_from_file {
    my ($self, $file) = @_;
    return unless $file =~ m/\.xls$/;

    if ($file =~ m/\.xls$/) {    # Process only xls file data.
        my $txtfile = tmpnam();
        _exceltotext($file, $txtfile);
        $file = $txtfile;
    }
    my $in;
    if (exists $self->{encoding}) {
        my $encoding = $self->{encoding}{data}[0];
        open $in, "<:encoding($encoding)", $file
            or die "Couldn't open $file: $!";
    } else {
        open $in, '<', $file
            or die "Couldn't open $file: $!";
    }
    while (<$in>) {
        $self->add_data('text' => 'UnStored' => $_);
    }
    unlink $file; #Remove the  text file, part of maintenance.
    return $self;
}

sub _exceltotext {
    ##This is the standard code taken from SpreadSheet::ParseExcel Module.
    my  $excel = shift;
    my  $output = shift;

    my $oExcel = Spreadsheet::ParseExcel->new();
    open my $txt_out, '>', $output or die "Not able to open file : $!";

    my $oBook = $oExcel->Parse($excel);
    my($iC, $oWkS, $oWkC);

    print $txt_out "FILE  :", $oBook->{File} , "\n";
    print $txt_out "COUNT :", $oBook->{SheetCount} , "\n";

    print $txt_out "AUTHOR:", $oBook->{Author} , "\n"
        if defined $oBook->{Author};

    for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++) {
        $oWkS = $oBook->{Worksheet}[$iSheet];
        print OUTPUT  $oWkS->{Name}, "\n";
        for(my $iR = $oWkS->{MinRow} ;
            defined $oWkS->{MaxRow} && $iR <= $oWkS->{MaxRow} ;
             $iR++)
        {
            for(my $iC = $oWkS->{MinCol} ;
                defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ;
                $iC++)
            {
                $oWkC = $oWkS->{Cells}[$iR][$iC];
                print OUTPUT $oWkC->Value, "\n" if($oWkC);
            }
        }
    }
    close($txt_out);
}

1;
