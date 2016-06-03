START:

use DBI;
use DBD::Oracle qw(:ora_types);
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use Text::CSV_XS;
#use DateKey_ARC;
use DBConnector;
use Win32::Job;
use Getopt::Long;
use IO::File;
use MIME::QuotedPrint;
use MIME::Base64;
use Mail::Sendmail;
use Date::Calc qw( Today Add_Delta_Days Month_to_Text);

($year,$month,$day) = Today();
$month_to_text = Month_to_Text($month);

$test = 1;
if ($test eq 1){

	
	&late_posted_sales_csv;
	&previous_day_sales;
	&insert_data;
	
	$date = qq{ 	SELECT TO_CHAR(SYSDATE-1,'DD-MM-YY') DATE_FLD FROM DUAL 	 }; 
	 
	my $sth_date_1 = $dbh->prepare ($date);
	 $sth_date_1->execute;

	while (my $x = $sth_date_1->fetchrow_hashref()) {
		$as_of = $x->{DATE_FLD};
	}
	
	
	exit;
	
}

else{
	print "Test Status = ". $test ." \nETL still running\n";
	
	$tst_query->finish();
	
	sleep(600);
	
	goto START;
}
 

sub late_posted_sales_csv{

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";
printf "Test ETL Status = ". $test ." \nArc BI Sales Performance Part 1\nGenerating Data from Source \n";
my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });


$test = qq{ 
	Select A.Store as STORE, Trunc(A.Tran_Datetime) As BUSINESS_DATE , To_Char(A.Tran_Datetime,'HH:MI:SS AM') As MY_TIME, 
             C.Cust_Id AS CUST_ID, C.Postal_Code AS POSTAL_ID, B.Tran_Seq_No AS TRAN_SEQ_NO, a.Tran_No AS TRAN_NO, B.Item_Seq_No AS ITEM_SEQ_NO, 
			 Substr(Lpad(a.Tran_No,10,0),1,3) As Register_Num, B.Item AS ITEM,
             B.Non_Merch_Item AS NON_MERCH_ITEM, B.Qty AS QTY, B.Unit_Retail AS UNIT_RETAIL, b.Total_Igtax_Amt AS TOTAL_IGTAX_AMT, 
			 E.Av_Cost  AS AV_COST From Sa_Tran_Head A Inner Join Sa_Tran_Item B On A.Store = B.Store
             AND A.TRAN_SEQ_NO = B.TRAN_SEQ_NO INNER JOIN SA_CUSTOMER C ON A.STORE = C.STORE AND A.DAY = C.DAY AND A.TRAN_SEQ_NO = C.TRAN_SEQ_NO
             INNER JOIN ITEM_LOC_SOH E ON B.ITEM = E.ITEM AND B.STORE = E.LOC WHERE 
             TRUNC(TRAN_DATETIME)= (SELECT TO_CHAR(SYSDATE-1,'DD MON YYYY') FROM DUAL) AND A.TRAN_TYPE = 'SALE' AND A.STATUS = 'P' AND C.CUST_ID IN 
             (SELECT CARD_NO_RESA FROM MG_MRC_ISO_ACCNTS)
			 };

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($sth);
 while (my $row = $sth->fetch) {
     $csv->print ($csv->error_diag);
     }
 

$dbh->disconnect; 

}

	
$query_handle->finish();

$workbook->close();
$dbh->disconnect; 



sub mail {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = 'lea.gonzaga@metroretail.com.ph';

$from = 'Report Mailer<report.mailer@metroretail.com.ph>';
		
$subject = 'ISO Sales ' . $as_of;

$msgbody_file = 'message_BI_LATE.txt';

$attachment_file = 'late_posted_'.$as_of .'.csv';

my $msgbody = read_file( $msgbody_file );

my $attachment_data = encode_base64( read_file( $attachment_file, 1 ));

my %mail = (
    To   => $to,
	From  => $from,
    Subject => $subject,
	'content-type' => "multipart/alternative; boundary=\"$boundary\""
);

$mail{smtp} = '10.190.1.30';
$mail{Cc} = $cc if $cc;
$mail{Bcc} = $bcc if $bcc;

my $boundary = "====" . time . "====";

$mail{'content-type'} = qq(multipart/mixed; boundary="$boundary");

$boundary = '--'.$boundary;

$mail{body} = 
<<END_OF_BODY;
$boundary
Content-Type: text/plain; charset="iso-8859-1"
Content-Transfer-Encoding: quoted-printable
$msgbody
$boundary
Content-Type: application/octet-stream; name="$attachment_file"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$attachment_file"
$attachment_data
$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

}

sub read_file {

my( $filename, $binmode ) = @_;
my $fh = new IO::File;
$fh->open("<".$filename) or die "Error opening $filename for reading - $!\n";
$fh->binmode if $binmode;
local $/;
<$fh>
	
}





