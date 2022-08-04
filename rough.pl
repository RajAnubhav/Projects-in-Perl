#!/usr/bin/perl
use Excel::Writer::XLSX;
use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Spreadsheet::ParseXLSX;
use Spreadsheet::XLSX;
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

    #check which first column is empty
    my $converter = Text::Iconv->new("utf-8", "windows-1251");
    my $excel = Spreadsheet::XLSX->new('atnd.xlsx', $converter);    #open the file
    my $sheet = (@{$excel->{Worksheet}})[0];                        #select the sheet
    my $col = ($sheet->{MaxCol})+1;                                 #column to take atendance in

    print "Enter the date : ";
    chomp($date = <STDIN>);
    $i=0;
    $writeTo=0;
    # printing into the excell sheet
    my $Excelbook = Excel::Writer::XLSX->new( './atnd.xls' );
    my $Excelsheet = $Excelbook->add_worksheet("Atnd Sheet");       #open the worksheet
    my $sheet =(@{$excel->{Worksheet}})[0];                         #select the worksheet
    $Excelsheet->write($writeTo++,0, $date );                       #write date as column heading  
    while(1){                                                       
        $sheet->{MaxCol} ||= $sheet->{MinCol};            
        foreach my $col (1 ..  $sheet->{MaxCol}) {                  #loop col 1 to max not 0 coz it has the titke "Name"          
            my $cell = $sheet->{Cells}[$row][0];                          
            if ($cell) {                                            #if cell is not empty
                printf(" %s (p/a): ", $cell->{Val});
                chomp(my $state = <STDIN>);
                $Excelsheet->write($writeTo++,$col, $firstName );   #write attendance
            }        
        }
    }
    $Excelbook->close();
}
sub atnd_count{
    # read and calculate atendance

    $nameRow=0;
    print "Enter name to print attendance: ";
    chomp($name = <STDIN>);                                         #name to find attendance for
    my $converter = Text::Iconv->new("utf-8", "windows-1251");
    my $excel = Spreadsheet::XLSX->new('atnd.xlsx', $converter);    #open file
    my $sheet = (@{$excel->{Worksheet}})[0];
    $sheet->{MaxRow} ||= $sheet->{MinRow};
    foreach my $row (1 .. $sheet->{MaxRow}) {                       #loop in row with the names
        my $cell = $sheet->{Cells}[$row][0];
        if($name == $cell){                                         #if name match found get column number
            $nameRow=$row;                          
            break;
        }
    }
    if($nameRow==0){
        print "Name not found in the records";
    }
    else{                                                           #if name is found
        $totalclass=0;
        $presentclass=0;
        $sheet->{MaxCol} ||= $sheet->{MinCol};            
        foreach my $col (1 ..  $sheet->{MaxCol}) {
            my $cell = $sheet->{Cells}[$nameRow][$col];
            if($cell){
                $totalclass++;
                if($cell=="p"){
                    $presentclass++;
                }
            }
        }
        print ("%s\n\tTotal no of days: %s \n\tNo of days present: %s\n",$name,$totalclass,$presentclass);
    }
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