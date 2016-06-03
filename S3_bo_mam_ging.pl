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


$test_query = qq{ SELECT CASE WHEN EXISTS (SELECT *
					FROM ADMIN_ETL_LOG 
					WHERE TO_DATE(LOG_DATE, 'DD-MON-YY') = TO_DATE(SYSDATE,'DD-MON-YY') AND TASK_ID = 'IntSalesDspTy' AND ERR_CODE = 0) THEN 1 ELSE 0 END STATUS 
					FROM DUAL };

$tst_query = $dbh->prepare($test_query);
$tst_query->execute();

while ( my $x =  $tst_query->fetchrow_hashref()){
	$test = $x->{STATUS};
} 
$test = 1;
if ($test eq 1){

	# $date = qq{ 
	# SELECT WEEK_NUMBER_THIS_YEAR, DATE_KEY, TO_CHAR(DATE_FLD, 'DD Mon YYYY') DATE_FLD, WEEK_ST_DATE_KEY, WEEK_END_DATE_KEY 
	# FROM DIM_DATE 
	# WHERE DATE_FLD = (SELECT TO_DATE(VALUE,'YYYY-MM-DD') FROM ADMIN_ETL_SUMMARY)
	 # };

	 #$date = qq{ 	SELECT TO_CHAR(SYSDATE-1,'DD MON YYYY') DATE_FLD FROM DUAL 	 }; 
	 $date = qq{ 	SELECT TO_CHAR(SYSDATE-1,'DD MON YYYY') DATE_FLD FROM DUAL 	 }; 
	 
	my $sth_date_1 = $dbh->prepare ($date);
	 $sth_date_1->execute;

	while (my $x = $sth_date_1->fetchrow_hashref()) {
		$as_of = $x->{DATE_FLD};
	}

	printf "Test ETL Status = ". $test ." \nArc BI Sales Performance Part 1\nGenerating Data from Source \n";
 
	$workbook = Excel::Writer::XLSX->new("Sales by Register (S13 as of $as_of).xlsx");
	$border1 = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 4 );
	$desc = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10 );
	$desc2 = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', text_wrap =>1, size => 10, shrink => 1 );
	
	&sale_header;
	
	$workbook->close();
	$tst_query->finish();
	$dbh->disconnect; 
	
	&mail;
	
	exit;
	
}

else{
	print "Test Status = ". $test ." \nETL still running\n";
	
	$tst_query->finish();
	
	sleep(600);
	
	goto START;
}
 
sub sale_header{

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
# my $uid = 'kent';
# my $pw = 'amer1c8';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $worksheet = $workbook->add_worksheet("sale_header");
	$worksheet->set_column( 0, 0, 5 );
	$worksheet->set_column( 1, 4, 12 );
	$worksheet->set_column( 5, 5, 12 );
	$worksheet->set_column( 6, 9, 12 );
	$worksheet->set_column( 10, 10, 12 );
	$worksheet->set_column( 11, 14, 12 );
	my $a = 2;
	my $col = 0;
	
	foreach my $i ( "DEPARTMENT", "DEPT_NAME", "TOT_SALES", "TAX", "SALES") {
		$worksheet->write( 1, $col++, $i, $desc2 );
	}
	
	my $query = qq {
	SELECT TABLE3.ID_DPT_POS AS DEPARTMENT, TABLE2.NM_DPT_POS AS DEPTNAME, 
SUM(TABLE1.TOTSALES) as totsales, SUM(TABLE1.TAX) as tax, SUM(TABLE1.TOTSALES)-SUM(TABLE1.TAX) AS SALES
FROM 
(SELECT TRN.DC_DY_BSN AS TRANDATE, TRN.ID_WS || TRN.AI_TRN AS VOIDED, RTN.ID_DPT_POS AS DEPARTMENT, RTN.ID_WS, RTN.AI_TRN, RTN.MO_EXTN_DSC_LN_ITM AS TOTSALES, RTN.ID_ITM AS ID_ITM, RTN.DE_ITM_SHRT_RCPT AS DE_ITM_SHRT_RCPT, (CASE WHEN RTN.FL_TX = '1' THEN ROUND((RTN.MO_EXTN_LN_ITM_RTN - (RTN.MO_EXTN_LN_ITM_RTN/1.12)),2) ELSE 0.00 END) AS TAX 
FROM TR_LTM_SLS_RTN RTN
INNER JOIN TR_TRN TRN ON RTN.DC_DY_BSN = TRN.DC_DY_BSN AND RTN.ID_WS = TRN.ID_WS AND RTN.AI_TRN = TRN.AI_TRN 
WHERE TRN.DC_DY_BSN BETWEEN '2015-05-01' AND '2015-05-31' AND TRN.TY_TRN IN (1,2) AND RTN.FL_VD_LN_ITM=0 AND TRN.SC_TRN = 2)TABLE1 
INNER JOIN AS_ITM TABLE3 ON TABLE1.ID_ITM = TABLE3.ID_ITM 
INNER JOIN ID_DPT_PS_I8 TABLE2 ON TABLE3.ID_DPT_POS = TABLE2.ID_DPT_POS 
WHERE TABLE1.VOIDED NOT IN (SELECT ID_WS_VD || AI_TRN_VD FROM TR_VD_PST WHERE DC_DY_BSN BETWEEN '2015-05-01' AND '2015-05-31')
GROUP BY TABLE3.ID_DPT_POS,TABLE2.NM_DPT_POS
ORDER BY TABLE3.ID_DPT_POS;
	};
	my $query_handle = $dbh->prepare($query);
	$query_handle->execute();
	
	while (my $y = $query_handle->fetchrow_hashref()) {
		$worksheet->write($a,0, $y->{DEPARTMENT},$desc);
		$worksheet->write($a,1, $y->{DEPTNAME},$desc);
		$worksheet->write($a,2, $y->{totsales},$desc);
		$worksheet->write($a,3, $y->{tax},$desc);
		$worksheet->write($a,4, $y->{SALES},$desc);
		#$worksheet->write($a,5, $y->{DESCRIPTION},$desc);
		#$worksheet->write($a,6, $y->{VENDOR_NAME},$desc);
		#$worksheet->write($a,7, $y->{COST},$border1);
		#$worksheet->write($a,8, $y->{RETAIL_PRICE},$border1);
		#$worksheet->write($a,9, $y->{QTY},$desc);
		#$worksheet->write($a,10, $y->{TOTAL_RETAIL_WITH_VAT},$border1);
		#$worksheet->write($a,11, $y->{CUR_INVEN},$desc);
		$a++;
	}
	
	$query_handle->finish();
	$dbh->disconnect; 

}


sub mail {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

#$to = 'april.agapito@metrogaisano.com, rosemarie.saravia@metrogaisano.com, rhoda.camporedondo@metrogaisano.com, jay.aguilar@metrogaisano.com, seth.jaminal@metrogaisano.com, miscellaneous.mercha@metrogaisano.com, teena.velasco@metrogaisano.com, richard.ordillo@metrogaisano.com';

$bcc = 'lea.gonzaga@metrogaisano.com';

#$bcc = ' rex.cabanilla@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'S3 Non Vat Sales (as of ' .$as_of. ') ';

$msgbody_file = 'message.txt';

$attachment_file = "S3 Non Vat Sales(as of $as_of).xlsx";

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





