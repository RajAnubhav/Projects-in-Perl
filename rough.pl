#!/usr/bin/perl
use Excel::Writer::XLSX;
use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel;

package Person;
sub new{
    my $class => shift;
    my $self = {
        _firstName => shift,
        _ssn => shift,
    };

    open(fh, ">>", "./Book1.xlsx");
    $a = "$self->{'_firstName'}$self->{'_ssn'}\n";
    print fh $a;
    close(fh) or "couldn't close the file";

    bless $self, $class;
    return $self;
}

my $ch = 1;
my @attend;

while($ch){
    print "Enter \n 1: for entering into datasheet\n 2: Check record\n 3: No. of days\n 4: for exiting\n";
    my $check = <STDIN>;

    if($check==1)
    # writing into the file
    {
        print "Enter the number of elements you want to enter : \n";
        $count = <STDIN>;
        $i=0;
        while($count > 0){
            $i++;
            print "Enter your name : \n";
            $firstName = <STDIN>;

            # printing into the excell sheet
            my $Excelbook = Excel::Writer::XLSX->new( './Book1.xlsx' );
            my $Excelsheet = $Excelbook->add_worksheet();

            # Writing values at A1 and A2
            $Excelsheet->write( "A$i", $firstName );
            $Excelsheet->write( "B$i", 1 );

            # -------------------------------
            $object1 = new Person ("$firstName", $ssn);    
            $count --;
        }
    }
    elsif($check == 2 ){
        # reading from the file
        open(fh, "./Book1.xlsx") or die "File '$filename' can't be opened\n";
        $i = 0;
        $firstline = <fh>;
        
        printf "The elements in the attendance sheet are : \n";
        while(<fh>)
        {
            
            print "$_";
            @attend[$i] = "$_";
            $i++;
        }
        close;
    }
    elsif($check == 3){
        print "welcome to section 3\n";
        $j=0;
        $c = 0;
        $count = 0;
            
        open(fh, "./Book1.xlsx") or die "File '$filename' can't be opened\n";
        my ($lines, $words, $chars) = (0,0,0);

        while (<fh>) {
            $lines++;
            $chars += length($_);
            $words += scalar(split(/\W+/, $_));
        }

        print("lines=$lines words=$words chars=$chars\n");
    }else{
        exit;
    }
}