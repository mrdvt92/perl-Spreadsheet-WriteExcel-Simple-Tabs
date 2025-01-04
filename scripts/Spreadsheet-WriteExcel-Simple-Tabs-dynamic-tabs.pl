#!/usr/bin/perl

=head1 NAME

Spreadsheet-WriteExcel-Simple-Tabs-dynamic-tabs.pl - Spreadsheet::WriteExcel::Simple::Tabs Dynamic Tabs from First Column

=cut

use strict;
use warnings;
use List::MoreUtils qw{uniq};
use Spreadsheet::WriteExcel::Simple::Tabs;
my $ss=Spreadsheet::WriteExcel::Simple::Tabs->new;

my @data=(
          ["tab",                        "Variable",                   "Value"               ],
          ["One",                        "Integer",                    "12345"               ],
          ["One",                        "Float",                      "12345.678"           ],
          ["Two",                        "Date MM/DD/YYYY HH24:MI:SS", "10/03/2010 23:15:30" ],
          ["Two",                        "Long Integer String",        "1234567890123456789" ],
          ["Three",                      "Zero Padded Number",         "02123456"            ],
         );

#Header
my $header=shift(@data); #remove header for later
shift(@$header);         #remove tab column name from header

#Clean up Tab Names
foreach my $row (@data) {
  $row->[0]=~s/[\[\]:\*\?\/\\]/ /g;                              #Invalid character []:*?/\ in worksheet name
  $row->[0]=substr($row->[0], 0, 31) if length($row->[0]) > 31;  #must be <= 31 chars
}

#Build List of Tabs
my @tabs=uniq(map {$_->[0]} @data); #preservers order

#Add Tab to Excel File
foreach my $tab (@tabs) {
  my @sheet=grep {$_->[0] eq $tab} @data; # grep for data on this tab
  shift @$_ foreach @sheet;               # remove the tab value from first column in sheet
  unshift @sheet, $header;                # add columns name for each sheet
  $ss->add($tab => \@sheet);              # add sheet to book
}

#print Excel file on STDOUT
print $ss->content;
