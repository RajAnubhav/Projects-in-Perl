#!/usr/bin/perl
use Excel::Writer::XLSX;
use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Spreadsheet::ParseXLSX;
use Switch::Plain;
sub input_names{
    # writing into the file
    print "Enter the number of names you want to enter : ";
    chomp($count = <STDIN>);
    $i=0;
    $writeTo=0;
    # printing into the excell sheet
    my $Excelbook = Excel::Writer::XLSX->new( './atnd.xlsx' );
    my $Excelsheet = $Excelbook->add_worksheet();
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
    my $sheet =(@{$excel->{Worksheet}})[0]; 
    my $writeStart=($sheet->{MaxRow})+1;

    print "Enter the number of names to insert : ";
    chomp($count = <STDIN>);
    $i=0;
    $writeTo=$writeStart;
    # printing into the excell sheet
    my $Excelbook = Excel::Writer::XLSX->new( './atnd.xlsx' );
    my $Excelsheet = ($Excelbook->sheets(0));
    while($count > 0){
        $i++;
        print "Enter your name $i : ";
        chomp ($firstName = <STDIN>);
        $Excelsheet->write($writeTo++,0, $firstName );    
        $count --;
    }
    $Excelbook->close();
    
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