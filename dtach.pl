#!/usr/bin/perl
########################################################
# AUTHOR = Michael Vincent
# www.VinsWorld.com
########################################################

use vars qw($VERSION);

$VERSION = "2.00 - 26 MAY 2015";

use strict;
use warnings;
use Getopt::Long qw(:config no_ignore_case);    #bundling
use Pod::Usage;

########################################################
# Start Additional USE
########################################################
use Cwd;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Outlook';
use Win32::OLE::Variant;
########################################################
# End Additional USE
########################################################

my %opt;
my ( $opt_help, $opt_man, $opt_versions );

GetOptions(
    'attach!'          => \$opt{attach},
    'A|appointments:s' => \$opt{dates},
    'conflict+'        => \$opt{conflict},
    'C|contacts:s'     => \$opt{contacts},
    'directory=s'      => \$opt{dir},
    'D|drafts:s'       => \$opt{drafts},
    'filter=s'         => \$opt{filter},
    'G|gal:s'          => \$opt{gal},
    'ignorecase!'      => \$opt{ignore},
    'I|inbox:s'        => \$opt{emails},
    'J|journal:s'      => \$opt{journal},
    'list!'            => \$opt{list},
    'N|notes:s'        => \$opt{notes},
    'output=s'         => \$opt{format},
    'O|outbox:s'       => \$opt{outbox},
    'prompt+'          => \$opt{prompt},
    'P|profile=s'      => \$opt{profile},
    'subfolder=s'      => \$opt{subfolder},
    'S|sentitems:s'    => \$opt{sentitems},
    'T|tasks:s'        => \$opt{tasks},
    'x|regex!'         => \$opt{regex},
    'X|deleteditems:s' => \$opt{deleteditems},
    'help!'            => \$opt_help,
    'man!'             => \$opt_man,
    'versions!'        => \$opt_versions
) or pod2usage( -verbose => 0 );

pod2usage( -verbose => 1 ) if defined $opt_help;
pod2usage( -verbose => 2 ) if defined $opt_man;

if ( defined $opt_versions ) {
    print
      "\nModules, Perl, OS, Program info:\n",
      "  $0\n",
      "  Version               $VERSION\n",
      "    strict              $strict::VERSION\n",
      "    warnings            $warnings::VERSION\n",
      "    Getopt::Long        $Getopt::Long::VERSION\n",
      "    Pod::Usage          $Pod::Usage::VERSION\n",
########################################################
# Start Additional USE
########################################################
      "    Cwd                 $Cwd::VERSION\n",
      "    Win32::OLE          $Win32::OLE::VERSION\n",
########################################################
# End Additional USE
########################################################
      "    Perl version        $]\n",
      "    Perl executable     $^X\n",
      "    OS                  $^O\n",
      "\n\n";
    exit;
}

########################################################
# Start Program
########################################################

# Get Outlook object
my $outlook = Win32::OLE->new('Outlook.Application');
die unless $outlook;
my $namespace = $outlook->GetNamespace("MAPI");

my ( $parent_mbox, $folder ) = GetProfile( $namespace, \%opt );

# subfolder provided - if not, defaults
if ( !defined $opt{subfolder} ) {
    if ( defined $opt{contacts} ) {
        $opt{subfolder} = "Contacts";
    } elsif ( defined $opt{gal} ) {
        $opt{subfolder} = "user";
    } elsif ( defined $opt{dates} ) {
        $opt{subfolder} = "Calendar";
    } elsif ( defined $opt{drafts} ) {
        $opt{subfolder} = "Drafts";
    } elsif ( defined $opt{outbox} ) {
        $opt{subfolder} = "Outbox";
    } elsif ( defined $opt{sentitems} ) {
        $opt{subfolder} = "Sent Items";
    } elsif ( defined $opt{deleteditems} ) {
        $opt{subfolder} = "Deleted Items";
    } elsif ( defined $opt{journal} ) {
        $opt{subfolder} = "Journal";
    } elsif ( defined $opt{notes} ) {
        $opt{subfolder} = "Notes";
    } elsif ( defined $opt{tasks} ) {
        $opt{subfolder} = "Tasks";
    } else {
        $opt{subfolder} = "Inbox";
    }
}

if ( ( defined $opt{subfolder} ) and ( defined $opt{gal} ) ) {
    if (    ( $opt{subfolder} !~ /^user$/i )
        and ( $opt{subfolder} !~ /^list$/i ) ) {
        print "$0: -G requires -s to be `user' or `list'\n";
        exit 1;
    }
}

# Just want a listing?
if ( $opt{list} ) {
    folder_list( $namespace, $folder, \%opt );
    exit;
}

if ( !defined $opt{attach} ) {
    $opt{attach} = 1;
}
$opt{conflict} = $opt{conflict} || 0;
$opt{prompt}   = $opt{prompt}   || 0;
$opt{ignore}   = $opt{ignore}   || 0;
$opt{regex}    = $opt{regex}    || 0;
$opt{dir}      = $opt{dir}      || cwd;

# Does it end with a \ ?  If not, add one
if ( $opt{dir} !~ /\\$/ ) {
    $opt{dir} = $opt{dir} . "\\";
}
# Does it start with a \ ?  If so, add the drive otherwise, it must
# be off the local directory, so add the current path
if ( $opt{dir} !~ /^[A-Za-z]:/ ) {
    if ( $opt{dir} =~ /^\\/ ) {
        my @drive = split( /:/, cwd );
        $opt{dir} = $drive[0] . ":" . $opt{dir};
    } else {
        $opt{dir} = cwd . "\\" . $opt{dir};
    }
}
# Replace all / with \ (must be \\ to escape the \)
$opt{dir} =~ s/\//\\/g;
# replace all \ with \\ so no errors later
$opt{dir} =~ s/\\/\\\\/g;
# Does directory exist?
if ( !( -e $opt{dir} ) ) {
    print "$0: directory not found - $opt{dir}\n";
    exit 1;
}

# Recursively search folders for the folder specified
my ( $subfolder, $items );
if ( defined $opt{gal} ) {
    $subfolder = $namespace->AddressLists;
    $items
      = $namespace->AddressLists->Item("Global Address List")->AddressEntries;
} else {
    $subfolder = loop_folders($folder);
    $items     = $subfolder->Items;
}

# If we found it, we're good to go
if ($subfolder) {
    print "Profile   : $parent_mbox\n";
    print "Subfolder : ", $opt{subfolder}, "\n";
    print "Items     : ", $items->Count, "\n\n";
} else {
    print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
    exit 1;
}

other_list( $subfolder, \%opt, \@ARGV );

########################################################
# End Program
########################################################

########################################################
# Begin Subroutines
########################################################

sub GetProfile {
    my ( $namespace, $args ) = @_;

    # If a folder is specified - find it
    if ( defined $args->{profile} ) {
        for ( 1 .. $namespace->Folders->Count ) {

            # Does it match the folder name provided in the argument
            if ( $namespace->Folders($_)->Name eq $args->{profile} ) {
                $parent_mbox = $namespace->Folders($_)->Name;
                $folder      = $namespace->Folders($_);
            }
        }

        # If we found it, we're good to go
        if ( !$parent_mbox ) {
            print "$0: folder not found - $args->{profile}\n";
            exit 1;
        }

    } else {
        if ( defined $args->{contacts} ) {
            $parent_mbox = $namespace->GetDefaultFolder(olFolderContacts)
              ->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderContacts);
        } elsif ( defined $args->{dates} ) {
            $parent_mbox = $namespace->GetDefaultFolder(olFolderCalendar)
              ->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderCalendar);
        } elsif ( defined $args->{drafts} ) {
            $parent_mbox
              = $namespace->GetDefaultFolder(olFolderDrafts)->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderDrafts);
        } elsif ( defined $args->{outbox} ) {
            $parent_mbox
              = $namespace->GetDefaultFolder(olFolderOutbox)->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderOutbox);
        } elsif ( defined $args->{sentitems} ) {
            $parent_mbox = $namespace->GetDefaultFolder(olFolderSentMail)
              ->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderSentMail);
        } elsif ( defined $args->{deleteditems} ) {
            $parent_mbox = $namespace->GetDefaultFolder(olFolderDeletedItems)
              ->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderDeletedItems);
        } elsif ( defined $args->{journal} ) {
            $parent_mbox
              = $namespace->GetDefaultFolder(olFolderJournal)->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderJournal);
        } elsif ( defined $args->{notes} ) {
            $parent_mbox
              = $namespace->GetDefaultFolder(olFolderNotes)->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderNotes);
        } elsif ( defined $args->{tasks} ) {
            $parent_mbox
              = $namespace->GetDefaultFolder(olFolderTasks)->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderTasks);
        } else {
            $parent_mbox
              = $namespace->GetDefaultFolder(olFolderInbox)->Parent->Name();
            $folder = $namespace->GetDefaultFolder(olFolderInbox);
        }
    }
    return ( $parent_mbox, $folder );
}

sub loop_folders {
    my $folder = shift;

    # Starts with \ must be fully qualified
    if ( $opt{subfolder} =~ /^\\/ ) {
        $opt{subfolder} =~ s/^\\//;
        my @subfolders = split /\\/, $opt{subfolder};
        if ( $opt{subfolder} ne $folder->Name ) {
            return;
        }
    }

    # Just relative path
    my @subfolders = split /\\/, $opt{subfolder};

    if ( ( $#subfolders == 0 ) && ( $subfolders[0] eq $folder->Name ) ) {
        return $folder;
    }

    for my $f ( 0 .. $#subfolders ) {
        my $MATCH = 0;
        if ( $folder->Folders->Count ) {
            for ( 1 .. $folder->Folders->Count ) {
### DEBUG:printf "%s = %s\n", $folder->Folders($_)->Name, $subfolders[$f];
                if ( $folder->Folders($_)->Name eq $subfolders[$f] ) {
                    if ( $f == $#subfolders ) {
                        return $folder->Folders($_);
                    } else {
                        $folder = $folder->Folders($_);
                        $MATCH  = 1;
                        last;
                    }
                }
            }
            if ( !$MATCH ) { return }
        }
    }
}

sub folder_list {
    my ( $namespace, $folder, $args ) = @_;

    # contacts fields
    if ( defined $args->{contacts} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->Items->Count ) > 0 ) {
                print "Contacts\n";
                Print_Opts( $folder, 1 );
            } else {
                print "$0: No Contacts found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

    } elsif ( defined $args->{gal} ) {
        print "Global Address List\n";
        print "  User\n";
        my $fields = getAbUserFields($namespace);
        for ( @{$fields} ) {
            print "    $_\n";
        }
        print "  Distribution List\n";
        $fields = getAbDLFields($namespace);
        for ( @{$fields} ) {
            print "    $_\n";
        }

        # calendar fields
    } elsif ( defined $args->{dates} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->{Items}->{Count} ) > 0 ) {
                print "Calendar\n";
                Print_Opts( $others, 1 );
            } else {
                print "$0: No Dates found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

        # drafts fields
    } elsif ( defined $args->{drafts} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->{Items}->{Count} ) > 0 ) {
                print "Drafts\n";
                Print_Opts( $others, 1 );
            } else {
                print "$0: No Drafts found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

        # outbox fields
    } elsif ( defined $args->{outbox} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->{Items}->{Count} ) > 0 ) {
                print "Outbox\n";
                Print_Opts( $others, 1 );
            } else {
                print "$0: No Emails found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

        # sent items fields
    } elsif ( defined $args->{sentitems} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->{Items}->{Count} ) > 0 ) {
                print "Sent Items\n";
                Print_Opts( $others, 1 );
            } else {
                print "$0: No Sent Items found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

        # deleted items fields
    } elsif ( defined $args->{deleteditems} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->{Items}->{Count} ) > 0 ) {
                print "Deleted Items\n";
                Print_Opts( $others, 1 );
            } else {
                print "$0: No Deleted Items found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

        # email fields
    } elsif ( defined $args->{emails} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->{Items}->{Count} ) > 0 ) {
                print "Inbox\n";
                Print_Opts( $others, 1 );
            } else {
                print "$0: No Emails found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

        # journal fields
    } elsif ( defined $args->{journal} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->{Items}->{Count} ) > 0 ) {
                print "Journal\n";
                Print_Opts( $others, 1 );
            } else {
                print "$0: No Journal entries found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

        # notes fields
    } elsif ( defined $args->{notes} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->{Items}->{Count} ) > 0 ) {
                print "Notes\n";
                Print_Opts( $others, 1 );
            } else {
                print "$0: No Notes found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

        # tasks fields
    } elsif ( defined $args->{tasks} ) {
        my $others = loop_folders($folder);
        if ($others) {
            if ( ( my $count = $others->{Items}->{Count} ) > 0 ) {
                print "Tasks\n";
                Print_Opts( $others, 1 );
            } else {
                print "$0: No Tasks found\n";
                exit 1;
            }
        } else {
            print "$0: subfolder not found - $parent_mbox\\$opt{subfolder}\n";
            exit 1;
        }

        # folder listing (DEFAULT)
    } else {
        for ( 1 .. $namespace->Folders->Count ) {

            if ( $namespace->Folders($_)->Name !~ /Public Folders/ ) {
                print "-------------\nProfile    : ",
                  $namespace->Folders($_)->Name, "\nSubfolders :\n";
                Print_Subs( $namespace->Folders($_), 1 );
            }
        }

        print "-------------\nProfile    : Address Books\nSubfolders :\n";
        for my $cnt ( 1 .. $namespace->AddressLists->Count ) {
            if ( $namespace->AddressLists->Item($cnt)->Name
                !~ /Public Folders/ ) {
                print "    "
                  . $namespace->AddressLists->Item($cnt)->Name . "\n";
            }
        }
    }

    # Enumerate [x] fields
    sub Print_Opts {
        my ( $others, $cnt ) = @_;

        for ( sort( keys( %{$others->{Items}->Item($cnt)} ) ) ) {
            print "  $_\n";
        }
    }

    # Enumerate folder listing
    sub Print_Subs {

        my $Top   = shift;
        my $depth = shift;

        #  Collection indexes start at 1, not 0
        for ( 1 .. $Top->Folders->Count ) {
            my $folder = $Top->Folders->Item($_)->Name;
            print "    " x ${depth}, "$folder\n";

            if ( $Top->Folders->Item($_)->Folders->Count > 0 ) {
                Print_Subs( $Top->Folders->Item($_), $depth + 1 );
            }
        }
    }
}

sub other_list {
    my ( $others, $args, $argv ) = @_;

    my $options;
    my $count;
    my $items;
    if ( defined $args->{gal} ) {
        $count = $others->Item("Global Address List")->AddressEntries->Count;
        $items = $others->Item("Global Address List");
    } else {
        $count = $others->Items->Count;
        $items = $others->Items;
    }
    my $filter = $args->{filter};

    # Contacts
    if ( defined $args->{contacts} ) {
        $options = $args->{contacts};
        if ( !defined $args->{filter} ) {
            $filter = 'FullName';
        }

        # GAL
    } elsif ( defined $args->{gal} ) {
        $options = $args->{gal};
        if ( !defined $args->{filter} ) {
            $filter = 'Name';
        }

        # Calendar
    } elsif ( defined $args->{dates} ) {
        $options = $args->{dates};
        if ( !defined $args->{filter} ) {
            $filter = 'Subject';
        }

        # Drafts
    } elsif ( defined $args->{drafts} ) {
        $options = $args->{drafts};
        if ( !defined $args->{filter} ) {
            $filter = 'To';
        }

        # Outbox
    } elsif ( defined $args->{outbox} ) {
        $options = $args->{outbox};
        if ( !defined $args->{filter} ) {
            $filter = 'To';
        }

        # Sent Items
    } elsif ( defined $args->{sentitems} ) {
        $options = $args->{sentitems};
        if ( !defined $args->{filter} ) {
            $filter = 'To';
        }

        # Deleted Items
    } elsif ( defined $args->{deleteditems} ) {
        $options = $args->{deleteditems};
        if ( !defined $args->{filter} ) {
            $filter = 'Subject';
        }

        # Journal
    } elsif ( defined $args->{journal} ) {
        $options = $args->{journal};
        if ( !defined $args->{filter} ) {
            $filter = 'Subject';
        }

        # Notes
    } elsif ( defined $args->{notes} ) {
        $options = $args->{notes};
        if ( !defined $args->{filter} ) {
            $filter = 'Subject';
        }

        # Tasks
    } elsif ( defined $args->{tasks} ) {
        $options = $args->{tasks};
        if ( !defined $args->{filter} ) {
            $filter = 'Subject';
        }

        # Emails (DEFAULT)
    } else {

        # This is default options so may not be specified as -I
        # Need to populate values if not specifically -I
        $options = $args->{emails} || '';
        $args->{emails} = $args->{emails} || '';
        if ( !defined $args->{filter} ) {
            $filter = 'SenderName';
        }
    }

    # if no items - done
    if ( $count == 0 ) { return }

    # Output format and file/STDOUT
    my $format = '';
    my $OUT    = \*STDOUT;
    my $OUTFILE;
    if ( defined $args->{format} ) {

        # CSV
        if ( $args->{format} =~ /^csv$/i ) {
            $format = 'csv'

              # output CSV file
        } elsif ( $args->{format} =~ /\.csv$/i ) {
            $format = 'csvfile';
            if ( !( open( $OUTFILE, '>', "$args->{format}" ) ) ) {
                print "$0: Cannot open output file - $args->{format}\n";
                exit 1;
            }
            $OUT = $OUTFILE

              # Tab-delimited
        } elsif ( ( $args->{format} =~ /^tab$/i )
            || ( $args->{format} =~ /^txt$/i ) ) {
            $format = 'tab'

              # output Tab-delimited file
        } elsif ( ( $args->{format} =~ /\.tab$/i )
            || ( $args->{format} =~ /\.txt$/i ) ) {
            $format = 'tabfile';
            if ( !( open( $OUTFILE, '>', "$args->{format}" ) ) ) {
                print "$0: Cannot open output file - $args->{format}\n";
                exit 1;
            }
            $OUT = $OUTFILE

              # DEFAULT:  list
        } elsif ( $args->{format} =~ /^list$/i ) {
            $format = 'list';
        }
    }

    # What values to get
    my @opts;
    if ( $options ne '' ) {
        @opts = split /,/, $options;
    }

    # If not specified, get them all
    if ( ( @opts == 0 ) && ( $format ne '' ) ) {
        if ( defined $args->{gal} ) {
            if ( $args->{subfolder} eq "user" ) {
                my $r = getAbUserFields($namespace);
                for ( @{$r} ) {
                    push @opts, $_;
                }
            } else {
                my $r = getAbDLFields($namespace);
                for ( @{$r} ) {
                    push @opts, $_;
                }
            }
        } else {
            for ( sort( keys( %{$items->Item($count)} ) ) ) {
                push @opts, $_;
            }
        }
    }

    # Loop emails, contacts, etc...
    for my $k ( 1 .. $count ) {
        my $other = $items->Item($k);

        if ( defined $args->{gal} ) {
            if ( $args->{subfolder} eq "user" ) {
                if ( $items->AddressEntries($k)->AddressEntryUserType == 0 ) {
                    $other = $items->AddressEntries($k)->GetExchangeUser;
                } else {
                    next;
                }
            } else {
                if ( $items->AddressEntries($k)->AddressEntryUserType == 1 ) {
                    $other = $items->AddressEntries($k)
                      ->GetExchangeDistributionList;
                } else {
                    next;
                }
            }
        }

        # Using search
        if ( @{$argv} > 0 ) {
            my $FOUND = 0;
            for my $name ( @{$argv} ) {
                if ( defined $other->{$filter} ) {

                    # REGEX
                    if ( $args->{regex} ) {

                        # ignore case
                        if ( $args->{ignore} ) {
                            $FOUND = 1 if ( $other->{$filter} =~ /$name/i );
                        } else {
                            $FOUND = 1 if ( $other->{$filter} =~ /$name/ );
                        }

                        # exact match
                    } else {

                        # ignore case
                        if ( $args->{ignore} ) {
                            $FOUND = 1
                              if ( lc( $other->{$filter} ) eq lc($name) );
                        } else {
                            $FOUND = 1 if ( $other->{$filter} eq $name );
                        }
                    }
                }
            }
            next if ( !$FOUND );
        }

        # Header
        if ( ( $k == 1 ) && ( $format =~ /file$/ ) ) {
            printf $OUT "%s\n", join ',', @opts;
        }

        # Data
        my $i = 0;
        for my $k (@opts) {

            # Defined values
            if ( defined $other->{$k} ) {
                if ( $format =~ /^csv/ ) {
                    print $OUT "," if ( $i++ > 0 );
                    print $OUT "$other->{$k}";
                } elsif ( $format =~ /^tab/ ) {
                    print $OUT "\t" if ( $i++ > 0 );
                    print $OUT "$other->{$k}";
                } else {
                    print $OUT "$k = $other->{$k}\n";
                }

                # no values
            } else {
                if ( $format =~ /^csv/ ) {
                    print $OUT "," if ( $i++ > 0 );
                    print $OUT '';
                } elsif ( $format =~ /^tab/ ) {
                    print $OUT "\t" if ( $i++ > 0 );
                    print $OUT '';
                }
            }
        }
        print $OUT "\n" if ( ( $format ne 'list' ) && ( $i > 0 ) );

        # Attachments
        # Want attachments
        if (    $args->{attach}
            and !defined( $args->{notes} )
            and !defined( $args->{gal} ) ) {
            my $attach = $other->Attachments();

            # are there any
            if ( $attach->Count > 0 ) {
                print "  Attachments = ", $attach->Count, "\n";

                # user want prompting per item
                if (( $args->{prompt} >= 2 )
                    && !get_answer(
                        "Examine " . $attach->Count . " attachments?"
                    )
                  ) {
                    print "    Skipping message - NO attachments saved!\n";
                    next;
                }

                for my $attach_index ( 1 .. $attach->Count ) {
                    my $attachment = $attach->item($attach_index);
                    my $filename   = $attachment->Filename;

                    # user want prompting per attachment
                    if ($args->{prompt}
                        && !get_answer(
                            "Save attachment:  \"" . $filename . "\""
                        )
                      ) {
                        print "    Attachement NOT saved!\n";
                        next;
                    }

                    sub get_answer {
                        my $prompt = shift;
                        my $uInput;
                        while (1) {
                            print "    $prompt [y/n]?";
                            chomp( my $uInput = <STDIN> );

                            if ( lc($uInput) eq "n" ) {
                                return 0;
                            }
                            if ( lc($uInput) eq "y" ) {
                                return 1;
                            }
                        }
                    }

                    my ($saveas) = $args->{dir} . $filename;

                    if ( -e $saveas ) {
                        print "    File EXISTS! - ";

                        # Default is to make a unique name
                        if ( $args->{conflict} == 0 ) {

                            # Get a unique filename by appending date
                            my @time = localtime();
                            $saveas
                              .= "."
                              . ( $time[5] + 1900 )
                              . ( ( ( $time[4] + 1 ) < 10 )
                                ? ( "0" . ( $time[4] + 1 ) )
                                : ( $time[4] + 1 ) )
                              . ( ( $time[3] < 10 ) ? ( "0" . $time[3] )
                                : $time[3] )
                              . ( ( $time[2] < 10 ) ? ( "0" . $time[2] )
                                : $time[2] )
                              . ( ( $time[1] < 10 ) ? ( "0" . $time[1] )
                                : $time[1] )
                              . ( ( $time[0] < 10 ) ? ( "0" . $time[0] )
                                : $time[0] );
                            print "Saving as:  $saveas\n";
                        }

                        # single -c means do nothing
                        if ( $args->{conflict} == 1 ) {
                            print "Ignoring:  $saveas\n";
                            next;
                        }

                        # multiple -c means overwrite
                        if ( $args->{conflict} >= 2 ) {
                            print "Overwriting:  $saveas\n";
                        }
                    }
                    print "    Saving:  $saveas\n";
                    $attachment->SaveAsFile($saveas);
                    if ( !-e $saveas ) {
                        print "$0: error saving attachment - $filename";
                    }
                }    # for (attachment)
            }    # if attach->count
        }    # if -a option
    }    # for items
    if ( $format =~ /file$/ ) {
        close($OUTFILE);
    }
}

sub getAbUserFields {
    my ($namespace) = @_;

    my @ret;
    for my $cnt ( 1 .. $namespace->AddressLists->Item("Global Address List")
        ->AddressEntries->Count ) {
        if ( $namespace->AddressLists->Item("Global Address List")
            ->AddressEntries($cnt)->AddressEntryUserType == 0 ) {
            for my $k (
                sort( keys(
                        %{  $namespace->AddressLists->Item(
                                "Global Address List")->AddressEntries($cnt)
                              ->GetExchangeUser
                        }
                ) )
              ) {
                push @ret, $k;
            }
            last;
        }
    }
    return \@ret;
}

sub getAbDLFields {
    my ($namespace) = @_;

    my @ret;
    for my $cnt ( 1 .. $namespace->AddressLists->Item("Global Address List")
        ->AddressEntries->Count ) {
        if ( $namespace->AddressLists->Item("Global Address List")
            ->AddressEntries($cnt)->AddressEntryUserType == 1 ) {
            for my $k (
                sort( keys(
                        %{  $namespace->AddressLists->Item(
                                "Global Address List")->AddressEntries($cnt)
                              ->GetExchangeDistributionList
                        }
                ) )
              ) {
                push @ret, $k;
            }
            last;
        }
    }
    return \@ret;
}

########################################################
# End Program
########################################################

__END__

########################################################
# Start POD
########################################################

=head1 NAME

DTACH - Detach and save Outlook attachments

=head1 SYNOPSIS

 dtach [options] [search [...]]

=head1 DESCRIPTION

Script saves Outlook attachments from emails in the Inbox (by default)
to the current directory (by default).  Also able to output contacts,
calendars, emails and tasks information and save attachments.

=head1 OPTIONS

The following options select the scope.

 -A [n[,n...]]  List calendar entries.  Optional comma separated list
 --appointments of fields prints only named fields.  Default is print
                no fields unless -o defined.

 -C [n[,n...]]  List contacts.  Optional comma separated list of fields
 --contacts     prints only named fields.  Default is print no fields
                unless -o defined.

 -D [n[,n...]]  List drafts.  Optional comma separated list of fields
 --drafts       prints only named fields.  Default is print no fields
                unless -o defined.

 -G [n[,n...]]  List Global Address List.  Optional comma separated list 
 --gal          of fields prints only named fields.  Default is print no 
                fields unless -o defined.

 -I [n[,n...]]  List emails.  Optional comma separated list of fields
 --inbox        prints only named fields.  Default is print no fields
                unless -o defined.
                NOTE:  This is the default behavior.

 -J [n[,n...]]  List journal entries.  Optional comma separated list
 --journal      of fields prints only named fields.  Default is print
                no fields unless -o defined.

 -N [n[,n...]]  List notes.  Optional comma separated list of fields
 --notes        prints only named fields.  Default is print no fields
                unless -o defined.

 -O [n[,n...]]  List outbox.  Optional comma separated list of fields
 --outbox       prints only named fields.  Default is print no fields
                unless -o defined.

 -S [n[,n...]]  List sent items.  Optional comma separated list of
 --sentitems    fields prints only named fields.  Default is print no
                fields unless -o defined.

 -T [n[,n...]]  List tasks.  Optional comma separated list of fields
 --tasks        prints only named fields.  Default is print no fields
                unless -o defined.

 -X [n[,n...]]  List deleted items.  Optional comma separated list of
 --deleteditems fields prints only named fields.  Default is print no
                fields unless -o defined.

The following options control the operation of the script.

 search         FullName of contact, Subject of calendar/task or From
                name in email to search on.  Use double-quotes to
                delimit if spaces - for example "Firstname Lastname".
                DEFAULT:  (or not specified) all.

 -a             Save attachments found.  Use --no-attach to not save.
 --attach       DEFAULT:  (or not specified) Save attachments.

 -c [-c]        Defines behavior if file to be saved already exists.
 --conflict     -c    = Ignore (don't save)
                -c -c = Overwrite existing file
                DEFAULT:  (or not specified) Create new unique filename.

 -d directory   Directory to save attachments to.
 --directory    DEFAULT:  (or not specified) Current directory.

 -f filter      Use filter as field to match 'search' against.
 --filter       DEFAULT:  (or not specified) FullName for contacts.
                                             Subject for calendar/tasks/
                                                         notes.
                                             SenderName for Inbox.
                                             To for Drafts/Outbox/Sent
                                                    Items

 -i             Ignore case of search string.
 --ignorecase   DEFAULT:  (or not specified) case-sensitive.

 -l             List all outlook folders and subfolders and exit.
 --list         Only local folders - public folders are excluded.
                List fields available for -C, -D, -I, -T if provided.

 -o format      Output format.  Valid options are:
 --output         list, csv, txt|tab, <filename>.csv, <filename>.txt
                DEFAULT:  (or not specified) list.

 -p [-p]        Prompt before saving any attachments found.
 --prompt       -p    = Prompt before saving each attachment.
                -p -p = Prompt to skip all attachments of a given
                        message.  If 'N', then prompt for each
                        individual attachment of given message as per
                        single -p option.
                DEFAULT:  (or not specified) Do not prompt before saving.

 -P profile     Outlook main folder/account that contains the subfolders
 --profile      to search items for attachments.  Case sensitive.
                DEFAULT:  (or not specified) Personal Folders
                          (main profile).

 -s folder      Outlook subfolder name to search emails for attachments.
 --subfolder    Case sensitive.
                In case of -G, valid values are `user' or `list' for User 
                and Distribution List respectively.
                DEFAULT:  (or not specified) Inbox.

 -x             Use search string as regular expression.
 --regex        DEFAULT:  (or not specified) exact match.

 --help         Print Options and Arguments.
 --man          Print complete man page.
 --versions     Print Modules, Perl, OS, Program info.

=head1 EXAMPLES

The following examples demonstrate some executions of this script.
This is not an all inclusive list.

By default, with no arguments, the script will search the Inbox
of the primary Outlook account and save all attachments found in
all emails.

=head2 FOLDER LISTING

To list all Outlook accounts and folders, use:

  dtach -l

=head2 FIELD LISTING

To list all possible fields for display from Contacts entries, use:

  dtach -C -l

If no Contacts are found in the main Contacts folder, but there are
contacts in a subfolder under the Contacts folder, use:

  dtach -C -s SubContactFolder -l

Where 'SubContactFolder' is the subfolder name.

For field listings of Appointments, Inbox or Tasks, uses -A, -I or -T
instead of -C, respectively.

=head2 CONTACTS

To save contacts' names, emails and phone numbers from the main Contacts
folder in the primary Outlook account to a CSV file and save any found
attachments to 'MyDirectory', use (note the following lines should all
be typed on the same command line before pressing "Enter/Return"):

  dtach -C FullName,Email1Address,BusinessTelephoneNumber
  -o contacts.csv -d MyDirectory

To only save the contact whose name is "John Doe", add "John Doe" to
the end of the above command line.

=head2 INBOX MESSAGES

To print email messages from the subfolder 'Sub' in the Outlook profile
'My Email', and not save attachments, use (note the following lines
should all be typed on the same command line before pressing
"Enter/Return"):

  dtach --no-attach -P "My Email" -s Sub -I
  ReceivedTime,SenderName,To,CC,Subject,Body

To only print messages that were received in May (provided the
ReceivedTime is reported as "5/09/2010 06:00:00 AM" for exmaple), add:

  -f ReceivedTime -x "^5/"

to the above command line.

=head1 LICENSE

This software is released under the same terms as Perl itself.
If you don't know what that means visit L<http://perl.com/>.

=head1 AUTHOR

Copyright (C) Michael Vincent 2008-2015

L<http://www.VinsWorld.com>

All rights reserved

=cut
