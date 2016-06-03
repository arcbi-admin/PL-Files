use strict;
use warnings;
use File::Spec::Functions qw(rel2abs);
use Win32::OLE;
use Win32::OLE::Variant;
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3;    # die on errors...

# get excel file name and full path
my $filename = shift;
$filename = rel2abs($filename);

# get PDF file name
my $pdffilename = $filename;
$pdffilename =~ s/\.xls.*/\.pdf/i;

# delete existing PDF file
if ( -e $pdffilename ) {
    print "pdf file exists already. removing file\n";
    unlink($pdffilename);
}


# Open document
# Create new MSExcel object and load constants
my $MSExcel = Win32::OLE->new( 'Excel.Application', 'Quit' )
  or die "Could not load MS Excel";
my $excel = Win32::OLE::Const->Load($MSExcel);
my $Book = $MSExcel->Workbooks->Open( { FileName => "$filename" } );

# Run Excel function "Save As ..."
$Book->ExportAsFixedFormat(
    {
        Type                 => xlTypePDF,
        Filename             => "$pdffilename",
        Quality              => $excel->{xlQualityStandard},
        IncludeDocProperties => $excel->{True},
        IgnorePrintAreas     => $excel->{False},
        OpenAfterPublish     => $excel->{False},
    }
);

# Close document
$Book->Close( { SaveChanges => $excel->{xlDoNotSaveChanges} } );

if ( -e $pdffilename ) {
    print "PDF file created\n";
} 

exit(0);