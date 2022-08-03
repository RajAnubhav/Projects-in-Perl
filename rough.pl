#!/usr/bin/perl
use Excel::Writer::XLSX;
use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel;
use Switch::Plain;
sub input_names{
    # writing into the file
    print "Enter the number of names you want to enter : ";
    chomp($count = <STDIN>);
    $i=0;
    $writeTo=0;
    # printing into the excell sheet
    my $Excelbook = Excel::Writer::XLSX->new( './atnd.xls' );
    my $Excelsheet = $Excelbook->add_worksheet("Atnd Sheet");
    $Excelsheet->write($writeTo++,0,"Names" );
    while($count > 0){
        $i++;
        print "Enter your name $i : ";
        chomp ($firstName = <STDIN>);
        # Writing values at A1 and A2
        $Excelsheet->write($writeTo++,0, $firstName );    
        $count --;
    }
    $Excelbook->close();
}
sub append_names{
    #to append new names
    use Text::Iconv;
my $converter = Text::Iconv->new("utf-8", "windows-1251");
 
# Text::Iconv is not really required.
# This can be any object with the convert method. Or nothing.
 
use Spreadsheet::XLSX;
 
my $excel = Spreadsheet::XLSX->new('atnd.xlsx', $converter);
 
foreach my $sheet (@{$excel->{Worksheet}}) {
 
    printf("Sheet: %s\n", $sheet->{Name});
     
    $sheet->{MaxRow} ||= $sheet->{MinRow};
     
    foreach my $row ($sheet->{MinRow} .. $sheet->{MaxRow}) {
          
        $sheet->{MaxCol} ||= $sheet->{MinCol};
         
        foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol}) {
         
            my $cell = $sheet->{Cells}[$row][$col];
     
            if ($cell) {
                printf("( %s , %s ) => %s\n", $row, $col, $cell->{Val});
            }
     
        }
     
    }
 
}
    # 
    # use Spreadsheet::ParseExcel;
    # my $parser   = Spreadsheet::ParseExcel->new();
    # my $workbook = $parser->parse('./atnd.xls');
 
    # if ( !defined $workbook ) {
    #     die $parser->error(), ".\n";
    # }
    
    # for my $worksheet ( $workbook->worksheets() ) {
    
    #     my ( $row_min, $row_max ) = $worksheet->row_range();
    #     my ( $col_min, $col_max ) = $worksheet->col_range();
    
    #     for my $row ( $row_min .. $row_max ) {
    #         for my $col ( $col_min .. $col_max ) {
    
    #             my $cell = $worksheet->get_cell( $row, $col );
    #             next unless $cell;
    
    #             print "Row, Col    = ($row, $col)\n";
    #             print "Value       = ", $cell->value(),       "\n";
    #             print "Unformatted = ", $cell->unformatted(), "\n";
    #             print "\n";
    #         }
    #     }
    # }
    # my $parser=Spreadsheet::ParseExcel->new();
    # my $Excelbook = $parser->parse('atnd.xlsx');
    # for my $worksheet (@{$Excelbook->worksheets() }) {
    #     my $cell = $worksheet->get_cell(0,0);
    #     print "$cell";
    # }
    # my $sheet = $Excelbook->worksheet(1);
    # print "$sheet";
}
sub take_atnd{
    #add attendance to the file
    print "Enter the date : ";
    chomp($date = <STDIN>);
    $i=0;
    $writeTo=0;
    # printing into the excell sheet
    my $Excelbook = Excel::Writer::XLSX->new( './atnd.xls' );
    my $Excelsheet = $Excelbook->add_worksheet("Atnd Sheet");
    $Excelsheet->write($writeTo++,0,"Names" );
    while($count > 0){
        $i++;
        print "Enter your name $i : ";
        chomp ($firstName = <STDIN>);
        # Writing values at A1 and A2
        $Excelsheet->write($writeTo++,0, $firstName );    
        $count --;
    }
    $Excelbook->close();
}
sub atnd_count{
    # read and calculate atendance
}

my $ch = 1;
my @attend;
while($ch){
    print "======================================================================\n";
    print "Enter \n1: enter name list\n2: append names\n3: take attendance\n4: atnd count\n5: exit\n";
    chomp(my $check = <STDIN>); 
    nswitch($check) {
        case 1: { input_names() }
        case 2: { append_names() }
        case 3: { take_atnd() }
        case 4: { atnd_count() }
        case 5: {
            print "Terminating...\n";
            print "======================================================================\n";
            exit 0; 
        }
        default:{ print "Error: invald option" }
    }
}