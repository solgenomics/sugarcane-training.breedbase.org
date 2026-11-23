#!/usr/bin/env perl

=head1

phenotype_properties_check.pl

=head1 SYNOPSIS

    $load_trait_props.pl -H [dbhost] -D [dbname] -I [input file] -o [ontology] -w

=head1 COMMAND-LINE OPTIONS

 -H  host name
 -D  database name
 -w  overwrite
 -t  Test run . Rolling back at the end.
 -o  Ontology name (from db table, e.g. "GO")
 -I  input file, either .xls or .xlsx format

=head2 DESCRIPTION

The input file should have the following column headers:

  trait_name
  trait_format
  trait_default_value
  trait_minimum
  trait_maximum
  trait_categories
  trait_details

=head2 AUTHOR

chris simoes (ccs263@cornell.edu)

March 2024

=cut

use strict;
use warnings;
use Getopt::Std;
use Data::Dumper;
use Bio::Chado::Schema;
use CXGN::Metadata::Schema;
use CXGN::Phenome::Schema;
use CXGN::DB::InsertDBH;
use CXGN::Trial;
use SGN::Model::Cvterm;
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

my $schema          = Bio::Chado::Schema->connect( sub { $dbh->get_actual_dbh() } );
my $metadata_schema = CXGN::Metadata::Schema->connect( sub { $dbh->get_actual_dbh() });
my $phenome_schema  = CXGN::Phenome::Schema->connect( sub { $dbh->get_actual_dbh() });

# Fetch all trial IDs
my @trial_ids = map { $_->project_id } $schema->resultset("Project::Project")->search({}, { columns => ['project_id'] });

# Process only first 5 trials for debugging
# my @selected_trial_ids = @trial_ids[0 .. ($#trial_ids < 4 ? $#trial_ids : 4)];

my (%trial_data, %trial_cols);

foreach my $trial_id (@trial_ids ) {
    my $t;
    eval { $t = CXGN::Trial->new({ bcs_schema => $schema, trial_id => $trial_id }); };
    if ($@ || !defined $t) {
        print STDERR "Skipping trial $trial_id (failed to initialize Trial object: $@)\n";
        next;
    }

    my $trial_name            = $t->get_name;
    my $location_name         = $t->get_location ? $t->get_location()->[1] : "Unknown";
    my $breeding_program_name = $t->get_breeding_program();
    my $year                 = $t->get_year();
    my $test_type            = $t->get_design_type();

    # Skip trials with no year or "Analise" design type
    if (!defined $year) {
        print STDERR "Skipping trial $trial_id (no associated year)\n";
        next;
    }
    if ($test_type eq "Analise") {
        print STDERR "Skipping trial $trial_id (design type: Analise)\n";
        next;
    }

    # Store trial details
    $trial_cols{$trial_id} = [ $year, $breeding_program_name, $trial_name, $location_name ];

    # Retrieve phenotype data
    my $traits = $t->get_traits_assayed();
    my @trait_ids = map { $_->[0] } @$traits;
    my $data = $t->get_stock_phenotypes_for_traits(\@trait_ids, 'all', ['plot_of', 'plant_of'], 'accession', 'subject');
    
    $trial_data{$trial_id} = $data;
}

# Organize phenotype observations
my (%obs, %traits);
foreach my $trial_id (keys %trial_data) {
    foreach my $line (@{$trial_data{$trial_id}}) {
        $obs{$trial_id}->{$line->[9]}->{$line->[1]}->{$line->[3]} = $line->[7];
        $traits{$line->[3]}++;
    }
}

# Fetch trait details from database
sub get_trait_details {
    my ($trait_id) = @_;
    my $trait_name = $trait_id;
    $trait_name =~ s/\|.*//;
    my $dbh = $schema->storage->dbh;

    # Correct SQL queries to fetch min/max values
    my $sql_min = "SELECT value FROM cvtermprop 
                   WHERE cvtermprop.cvterm_id = (SELECT cvterm_id FROM cvterm WHERE name = ?) 
                   AND cvtermprop.type_id = (SELECT cvterm_id FROM cvterm WHERE name = 'trait_minimum')";

    my $sql_max = "SELECT value FROM cvtermprop 
                   WHERE cvtermprop.cvterm_id = (SELECT cvterm_id FROM cvterm WHERE name = ?) 
                   AND cvtermprop.type_id = (SELECT cvterm_id FROM cvterm WHERE name = 'trait_maximum')";

    my $sql_name = "SELECT name FROM cvterm WHERE name = ?";

    my $trait_min = $dbh->selectrow_array($sql_min, undef, $trait_name);
    my $trait_max = $dbh->selectrow_array($sql_max, undef, $trait_name);
    
    # Remove everything after "|" in trait name
    

    # If both values are missing, return undef to indicate skipping
    return undef unless defined $trait_min || defined $trait_max;

    return ($trait_min, $trait_max, $trait_name);
}


# Identify outliers and store them for the QC report
my @qc_issues;
foreach my $trial_id (keys %obs) {
    my ($year, $breeding_program_name, $trial_name, $location_name) = @{$trial_cols{$trial_id}};

    foreach my $accession (keys %{$obs{$trial_id}}) {
        foreach my $plot (keys %{$obs{$trial_id}->{$accession}}) {
            foreach my $trait_id (keys %{$obs{$trial_id}->{$accession}->{$plot}}) {


                # Attempt to fetch trait details, skip on failure
                
                my ($trait_min_value, $trait_max_value, $trait_name);
                eval {
                    ($trait_min_value, $trait_max_value, $trait_name) = get_trait_details($trait_id);
                };

                if ($@ || !defined $trait_name) {
                    print STDERR "Skipping trait $trait_id (failed to retrieve details: $@)\n";
                    next;
                }

                # Skip traits without minimum and maximum
                next unless defined $trait_min_value || defined $trait_max_value;

                my $observed_value = $obs{$trial_id}->{$accession}->{$plot}->{$trait_id};
                my $qc_type;

                if (defined $trait_min_value && defined $observed_value && $observed_value < $trait_min_value) {
                    $qc_type = "below minimum";
                }
                elsif (defined $trait_max_value && defined $observed_value && $observed_value > $trait_max_value) {
                    $qc_type = "above maximum";
                }

                if ($qc_type) {
                    push @qc_issues, {
                        year       => $year,
                        breeding_program_name => $breeding_program_name,
                        trial_name => $trial_name,
                        trial_location => $location_name,
                        accession  => $accession,
                        plot       => $plot,
                        trait_name => $trait_name,
                        observed   => $observed_value,
                        minimum    => $trait_min_value // 'N/A',
                        maximum    => $trait_max_value // 'N/A',
                        type       => $qc_type
                    };
                }
            }
        }
    }
}


# Create Excel file
my $date = strftime("%Y-%m-%d", localtime);
my $workbook = Excel::Writer::XLSX->new("/home/production/data_exchange/pheno_qc_$date.xlsx");
my $worksheet = $workbook->add_worksheet("QC Results");
$worksheet->write_row(0, 0, [qw(Year Breeding_Program Trial_Name Trial_Location Plot Accession Trait_Name Observed Minimum Maximum Type)]);

my $row = 1;
foreach my $entry (@qc_issues) {
    $worksheet->write_row($row++, 0, [@$entry{qw(year breeding_program_name trial_name trial_location plot accession trait_name observed minimum maximum type)}]);
}

$workbook->close();
print STDERR "QC results saved!\n";
