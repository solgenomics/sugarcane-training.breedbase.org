#!/usr/bin/env perl

=head1

stock_name_check.pl

=head1 SYNOPSIS

    $stock_name_check.pl -H [dbhost] -D [dbname] -I [input file] -o [ontology] -w

=head1 COMMAND-LINE OPTIONS

 -H  host name
 -D  database name
 -w  overwrite
 -t  Test run . Rolling back at the end.
 -o  Ontology name (from db table, e.g. "GO")
 -I  input file, either .xls or .xlsx format

=head2 DESCRIPTION

The input file should have the following column headers:


=head2 AUTHOR

chris simoes (ccs263@cornell.edu)

March 2024

=cut

use strict;
use warnings;
use Getopt::Std;
use CXGN::DB::InsertDBH;
use Bio::Chado::Schema;
use Excel::Writer::XLSX;
use POSIX qw(strftime);

# Parse command-line options
our ($opt_H, $opt_D, $opt_U, $opt_P);
getopts('H:D:U:P');

my $dbhost = $opt_H;
my $dbname = $opt_D;
my $dbuser = $opt_U;
my $dbpass = $opt_P;

# Database connection
print STDERR "Connecting to database...\n";
my $dbh = CXGN::DB::InsertDBH->new({
    dbhost   => $dbhost,
    dbname   => $dbname,
    dbargs   => { AutoCommit => 0, RaiseError => 1 }
});

my $schema = Bio::Chado::Schema->connect( sub { $dbh->get_actual_dbh() });

# Fetch the cvterm_id for "accession"
my $acc_cvterm = $schema->resultset("Cv::Cvterm")->find(
    { name => 'accession' }
);

if (!$acc_cvterm) {
    die "Error: Could not find cvterm for 'accession'\n";
}

my $acc_cvterm_id = $acc_cvterm->cvterm_id;

# Prepare Excel file
my $date = strftime("%Y-%m-%d", localtime);
my $excel_file = "/data_exchange/accession_validation_$date.xlsx";
my $workbook = Excel::Writer::XLSX->new($excel_file);
my $worksheet = $workbook->add_worksheet();

# Add headers to the Excel sheet
$worksheet->write(0, 0, "Accession Name");
$worksheet->write(0, 1, "Validation Message");

# Fetch all stocks of type "accession"
my $stock_rs = $schema->resultset("Stock::Stock")->search(
    { type_id => $acc_cvterm_id }
);

# Process each stock and validate
my $row = 1;
while (my $stock = $stock_rs->next) {
    my $uniquename = $stock->uniquename;
    
    print "Validating: $uniquename\n";
    
    my $validation_message = validate_string($uniquename);
    
    # Save to Excel file
    $worksheet->write($row, 0, $uniquename);
    $worksheet->write($row, 1, $validation_message);
    
    $row++;
}

# Close Excel workbook
$workbook->close();
print "Validation results saved to $excel_file\n";

# Subroutine to validate accession names
sub validate_string {
    my ($input) = @_;
    my @errors;

    # General rule: Minimum of 3 characters
    if (length($input) < 3) {
        push @errors, "Error: The string must have at least 3 characters.";
    }

    # General rule: Only uppercase letters and numbers allowed (except permitted special characters)
    if ($input =~ /[^A-Z0-9_\-\.\+\:\@\)\/\[\]]/){
        push @errors, "Error: Invalid characters detected. Only numbers, uppercase letters, and allowed special characters are permitted.";
    }


    # Limited occurrence rules for special characters
    my %char_limits = (
        ')' => 1, '[' => 1, '/' => 1,
        '.' => 2, '+' => 2, ':' => 2, '@' => 2
    );

    foreach my $char (keys %char_limits) {
        my $count = () = $input =~ /\Q$char\E/g;
        if ($count > $char_limits{$char}) {
            push @errors, "Error: The character '$char' appears $count times, but the maximum allowed is $char_limits{$char}.";
        }
    }

    # Special rule: "@" can only be followed by another "@" or ")"
    if ($input =~ /@[^@\)0-9]/) {
        push @errors, "Error: The character '\@' can only be followed by '\@', ')', or a number.";
    }

    # Specific prefix rules

    # Hybrid naming rule for productivity trials
    if ($input =~ /^(SX|SA)(\d{2})([MHPNT])?(\d{4})?([A-Z0-9]*)?$/) {
        my $num = $2;
        if ($num < 24) {
            push @errors, "Error: The third and fourth characters must be a number greater than 24.";
        }
        if ($5 && $5 =~ /[^A-Z0-9]/) {
            push @errors, "Error: After the ninth character, only uppercase letters and numbers are allowed.";
        }
    }

    # Hybrid naming rule for TopCross
    elsif ($input =~ /^(CL|CH|SB|SU)(\d{2})([A-Z])(\d{5})?([A-Z0-9]*)?$/) {
        my $num = $2;
        if ($num < 24) {
            push @errors, "Error: The third and fourth characters must be a number greater than 24.";
        }
        if ($5 && $5 =~ /[^A-Z0-9]/) {
            push @errors, "Error: After the tenth character, only uppercase letters and numbers are allowed.";
        }
    }

    # Classic breeding population naming rule
    elsif ($input =~ /^(\d{2})([ABCD])(SHE|UDI|TLD)(_)([A-Z0-9]*)?$/) {
        my $num = $1;
        if ($num < 24) {
            push @errors, "Error: The first two characters must be a number greater than 24.";
        }
    }

    # Return validation result
    if (@errors) {
        return join("; ", @errors);
    } else {
        return "The string is valid!";
    }
}