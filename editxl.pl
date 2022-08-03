#!/usr/bin/perl
 
use strict;
use warnings;
 
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
 
 
# Open an existing file with SaveParser
my $parser   = Spreadsheet::ParseExcel::SaveParser->new();
my $template = $parser->Parse('./atnd.xlsx');
 
 
# Get the first worksheet.
my $worksheet = $template->worksheet(0);
my $row  = 0;
my $col  = 0;
 
 
# Overwrite the string in cell A1
$worksheet->AddCell( $row, $col, 'New string' );
 
 
# Add a new string in cell B1
$worksheet->AddCell( $row, $col + 1, 'Newer' );
 
 
# Add a new string in cell C1 with the format from cell A3.
my $cell = $worksheet->get_cell( $row + 2, $col );
my $format_number = $cell->{FormatNo};
 
$worksheet->AddCell( $row, $col + 2, 'Newest', $format_number );
 
 
# Write over the existing file or write a new file.
$template->SaveAs('newfile.xlsx');