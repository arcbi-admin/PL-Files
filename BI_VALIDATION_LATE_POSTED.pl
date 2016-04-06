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

# $test_query = qq{ SELECT CASE WHEN EXISTS (SELECT SEQ_NO, ETL_SUMMARY, VALUE, ARC_DATE FROM ADMIN_ETL_SUMMARY WHERE TO_DATE(ARC_DATE, 'DD-MON-YY') = TO_DATE(SYSDATE,'DD-MON-YY')) THEN 1 ELSE 0 END STATUS FROM DUAL };

# $tst_query = $dbh->prepare($test_query);
# $tst_query->execute();

# while ( my $x =  $tst_query->fetchrow_hashref()){
	# $test = $x->{STATUS};
# } 
$test = 1;
if ($test eq 1){

	# $date = qq{ 
	# SELECT WEEK_NUMBER_THIS_YEAR, DATE_KEY, TO_CHAR(DATE_FLD, 'DD Mon YYYY') DATE_FLD, WEEK_ST_DATE_KEY, WEEK_END_DATE_KEY 
	# FROM DIM_DATE 
	# WHERE DATE_FLD = (SELECT TO_DATE(VALUE,'YYYY-MM-DD') FROM ADMIN_ETL_SUMMARY)
	 # };

	# my $sth_date_1 = $dbh->prepare ($date);
	 # $sth_date_1->execute;

	# while (my $x = $sth_date_1->fetchrow_hashref()) {
		# $as_of = $x->{DATE_FLD};
	# }

	# printf "Test ETL Status = ". $test ." \nArc BI Sales Performance Part 1\nGenerating Data from Source \n";
	
	&late_posted_sales_csv;
	&previous_day_sales;
	&insert_data;
	# $tst_query->finish();
	
	#&mail;
	
	# $workbook = Excel::Writer::XLSX->new('Bypass('.$day . '-' . $month_to_text . '-' .$year . ').xlsx');
	# $border1 = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 4 );
	# $desc = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10 );
	# $desc2 = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', text_wrap =>1, size => 10, shrink => 1 );
	
	# &bypass;
	
	# $workbook->close();
	
	exit;
	
}

else{
	print "Test Status = ". $test ." \nETL still running\n";
	
	$tst_query->finish();
	
	sleep(600);
	
	goto START;
}
 

sub late_posted_sales{

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'kent';
my $pw = 'amer1c8';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

$workbook = Excel::Writer::XLSX->new('ARC BI VALIDATION - Late Posted ('.$day . '-' . $month_to_text . '-' .$year . ').xlsx');
$border1 = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 4 );
$desc = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10 );
$desc2 = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', text_wrap =>1, size => 10, shrink => 1 );

my $worksheet = $workbook->add_worksheet("late_posted_sales");
$worksheet->set_column( 0, 0, 5 );
$worksheet->set_column( 1, 1, 12 );
$worksheet->set_column( 2, 3, 10 );
$worksheet->set_column( 4, 4, 12 );
my $a = 2;
my $col = 0;

foreach my $i ( "STORE", "TRAN_SEQ_NO", "TRAN_DATETIME", "UPDATE_DATETIME", "VALUE" ) {
	$worksheet->write( 1, $col++, $i, $desc2 );
}
	
my $query = qq {
SELECT STORE, TRAN_SEQ_NO, TO_CHAR(TRAN_DATETIME, 'DD-MON-YY HH:MM:SS') TRAN_DATETIME, TO_CHAR(UPDATE_DATETIME, 'DD-MON-YY HH:MM:SS') UPDATE_DATETIME, VALUE
FROM SA_TRAN_HEAD 
WHERE SUB_TRAN_TYPE IN ('SALE','LAYCMP','RETURN') AND STATUS = 'P' AND
TRUNC(TRAN_DATETIME) = (SELECT TO_CHAR(SYSDATE-1,'DD-MON-YY') FROM DUAL) AND TRUNC(UPDATE_DATETIME) = (SELECT TO_CHAR(SYSDATE,'DD-MON-YY') FROM DUAL)
};

my $query_handle = $dbh->prepare($query);
$query_handle->execute();

while (my $y = $query_handle->fetchrow_hashref()) {
	$worksheet->write($a,0, $y->{STORE},$desc);
	$worksheet->write($a,1, $y->{TRAN_SEQ_NO},$border1);
	$worksheet->write($a,2, $y->{TRAN_DATETIME},$border1);
	$worksheet->write($a,3, $y->{UPDATE_DATETIME},$border1);
	$worksheet->write($a,4, $y->{VALUE},$border1);
	$a++;
}
	
$query_handle->finish();

$workbook->close();
$dbh->disconnect; 

}

sub late_posted_sales_csv{

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
open my $fh, ">", 'late_posted_'.$day . '_' . $month_to_text . '_' .$year . '.csv' or die 'late_posted_'.$day . '_' . $month_to_text . '_' .$year . '.csv: $!';

$test = qq{ 
SELECT SALE.STORE, SALE.STORE_NAME, SALE.MERCH_GROUP_CODE_REV, SALE.GROUP_NO, ROUND(SUM((NVL(SALE.ACTUAL_AMT,0))-(NVL(SALE.TAX_AMT,0))-(NVL(SALE.DISC_AMT,0))),2) VALUE /*, SUM(((nvl(SALE.ACTUAL_AMT,0))-((nvl(SALE.ACTUAL_AMT,0))*((nvl(SALE.IGTAX_RATE,0))/(100 + (nvl(SALE.IGTAX_RATE,0))))))-((nvl(SALE.DISC_AMT,0))-((nvl(SALE.DISC_AMT,0))*((nvl(SALE.IGTAX_RATE,0))/(100 + (nvl(SALE.IGTAX_RATE,0))))))) AS NET_SALES */	
	FROM
	(SELECT TRUNC(H.TRAN_DATETIME) TRANS_DATE, H.STORE, S.STORE_NAME, H.TRAN_SEQ_NO, H.TRAN_NO, I.ITEM_SEQ_NO, 
	CASE WHEN (B.MERCH_GROUP_CODE = 'OT' AND I.DIVISION = 9000) THEN 'DS'     WHEN (B.MERCH_GROUP_CODE = 'OT' AND I.DIVISION = 8500) THEN 'SU'     WHEN (B.MERCH_GROUP_CODE = 'OT' AND I.DIVISION = 8000 AND G.GROUP_NO = 8040) THEN 'SU'     WHEN (B.MERCH_GROUP_CODE = 'OT' AND I.DIVISION = 8000 AND G.GROUP_NO != 8040) THEN 'DS'ELSE B.MERCH_GROUP_CODE END AS MERCH_GROUP_CODE_REV,
	G.GROUP_NO, I.DEPT, I.ITEM, case when i.ref_no5 in(0707, 0709) then 0 else TAX.IGTAX_RATE end as IGTAX_RATE, SUM(I.QTY) SALE_QTY, SUM(I.QTY*I.UNIT_RETAIL) ACTUAL_AMT, SUM(I.TOTAL_IGTAX_AMT) TAX_AMT, SUM(TAX.TOTAL_IGTAX_AMT) TAX_TAX, SUM(DISC.DISC_AMT) DISC_AMT
	FROM SA_TRAN_HEAD H 
		INNER JOIN SA_TRAN_ITEM I ON H.STORE = I.STORE AND H.TRAN_SEQ_NO = I.TRAN_SEQ_NO
		LEFT JOIN (SELECT STORE, TRAN_SEQ_NO, ITEM_SEQ_NO, DISC_TYPE, SUM(QTY*UNIT_DISCOUNT_AMT) DISC_AMT FROM SA_TRAN_DISC GROUP BY STORE, TRAN_SEQ_NO, ITEM_SEQ_NO, DISC_TYPE)DISC ON H.STORE = DISC.STORE AND I.TRAN_SEQ_NO = DISC.TRAN_SEQ_NO AND I.ITEM_SEQ_NO = DISC.ITEM_SEQ_NO
		LEFT JOIN SA_TRAN_IGTAX TAX ON H.STORE = TAX.STORE AND H.TRAN_SEQ_NO = TAX.TRAN_SEQ_NO AND I.ITEM_SEQ_NO=TAX.ITEM_SEQ_NO
		INNER JOIN STORE S ON H.STORE = S.STORE
		INNER JOIN DEPS ON I.DEPT = DEPS.DEPT
        INNER JOIN GROUPS G ON DEPS.GROUP_NO = G.GROUP_NO
		INNER JOIN DIVISION I ON G.DIVISION = I.DIVISION 
		INNER JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION
	WHERE (TRUNC(H.TRAN_DATETIME) = (SELECT TO_CHAR(SYSDATE-1,'DD-MON-YY') FROM DUAL) AND TRUNC(H.UPDATE_DATETIME) = (SELECT TO_CHAR(SYSDATE,'DD-MON-YY') FROM DUAL)) 
		AND H.SUB_TRAN_TYPE IN ('SALE','LAYCMP','RETURN') AND H.STATUS = 'P' 
	GROUP BY TRAN_DATETIME, H.STORE, S.STORE_NAME, H.TRAN_SEQ_NO, H.TRAN_NO, I.ITEM_SEQ_NO, 
	CASE WHEN (B.MERCH_GROUP_CODE = 'OT' AND I.DIVISION = 9000) THEN 'DS'     WHEN (B.MERCH_GROUP_CODE = 'OT' AND I.DIVISION = 8500) THEN 'SU'     WHEN (B.MERCH_GROUP_CODE = 'OT' AND I.DIVISION = 8000 AND G.GROUP_NO = 8040) THEN 'SU'     WHEN (B.MERCH_GROUP_CODE = 'OT' AND I.DIVISION = 8000 AND G.GROUP_NO != 8040) THEN 'DS'ELSE B.MERCH_GROUP_CODE END,
	G.GROUP_NO, I.DEPT, I.ITEM, CASE WHEN I.REF_NO5 IN(0707, 0709) THEN 0 ELSE TAX.IGTAX_RATE END)SALE
GROUP BY SALE.STORE, SALE.STORE_NAME, SALE.MERCH_GROUP_CODE_REV, SALE.GROUP_NO
ORDER BY 1, 3
};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_uc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die 'late_posted_'.$day . '_' . $month_to_text . '_' .$year . '.csv: $!';

$dbh->disconnect; 

}

sub previous_day_sales{

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
open my $fh, ">", 'previous_day_sales_'.$day . '_' . $month_to_text . '_' .$year . '.csv' or die 'previous_day_sales_'.$day . '_' . $month_to_text . '_' .$year . '.csv: $!';

$test = qq{ 
SELECT STORE, GROUP_NO, 
		SUM(SALE_QTY) SALE_QTY, 
		SUM(TAX_TAX) TAX_AMT, 
		SUM(NVL(DISC_AMT,0)) DISC_AMT, 
		SUM(NVL(ACTUAL_AMT,0)) GROSSSALE,
		ROUND(SUM(((NVL(ACTUAL_AMT,0))-((NVL(ACTUAL_AMT,0))*((NVL(IGTAX_RATE,0))/(100 + (NVL(IGTAX_RATE,0))))))-((NVL(DISC_AMT,0))-((NVL(DISC_AMT,0))*((NVL(IGTAX_RATE,0))/(100 + (NVL(IGTAX_RATE,0))))))),2) AS NET_SALE 	
		FROM
			(SELECT TRUNC(H.TRAN_DATETIME) TRANS_DATE, H.STORE, H.TRAN_SEQ_NO, H.TRAN_NO, I.ITEM_SEQ_NO, DEPS.GROUP_NO, I.DEPT, I.CLASS, I.SUBCLASS, I.ITEM, MST.SHORT_DESC, case when i.ref_no5 in(0707, 0709) then 0 else TAX.IGTAX_RATE end as IGTAX_RATE, SUM(I.QTY) SALE_QTY, SUM(I.QTY*I.UNIT_RETAIL) ACTUAL_AMT, SUM(I.TOTAL_IGTAX_AMT) TAX_AMT, SUM(TAX.TOTAL_IGTAX_AMT) TAX_TAX, SUM(DISC.DISC_AMT) DISC_AMT
			FROM SA_TRAN_HEAD H 
				JOIN SA_TRAN_ITEM I ON H.STORE = I.STORE AND H.TRAN_SEQ_NO = I.TRAN_SEQ_NO
				LEFT JOIN SA_TRAN_IGTAX TAX ON H.STORE = TAX.STORE AND H.TRAN_SEQ_NO = TAX.TRAN_SEQ_NO AND I.ITEM_SEQ_NO=TAX.ITEM_SEQ_NO
				LEFT JOIN (SELECT STORE, TRAN_SEQ_NO, ITEM_SEQ_NO, SUM(QTY*UNIT_DISCOUNT_AMT) DISC_AMT FROM SA_TRAN_DISC GROUP BY STORE, TRAN_SEQ_NO, ITEM_SEQ_NO)DISC ON H.STORE = DISC.STORE AND H.TRAN_SEQ_NO = DISC.TRAN_SEQ_NO AND I.ITEM_SEQ_NO = DISC.ITEM_SEQ_NO
				LEFT JOIN ITEM_MASTER MST ON I.ITEM = MST.ITEM
				LEFT JOIN DEPS ON MST.DEPT = DEPS.DEPT
			WHERE TRUNC(H.TRAN_DATETIME) = (SELECT TRUNC(SYSDATE-1) FROM DUAL) AND H.SUB_TRAN_TYPE IN ('SALE','LAYCMP','RETURN') AND H.STATUS = 'P'
			GROUP BY TRAN_DATETIME, H.STORE, H.TRAN_SEQ_NO, H.TRAN_NO, I.ITEM_SEQ_NO, DEPS.GROUP_NO, I.DEPT, I.CLASS, I.SUBCLASS, I.ITEM, MST.SHORT_DESC, CASE WHEN I.REF_NO5 IN(0707, 0709) THEN 0 ELSE TAX.IGTAX_RATE END)
GROUP BY STORE, GROUP_NO ORDER BY 1, 2
};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_uc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die 'previous_day_sales_'.$day . '_' . $month_to_text . '_' .$year . '.csv: $!';

$dbh->disconnect; 

}

sub bypass{

my $hostname = "10.128.0.220";
my $sid = "METROBIP";
my $port = '1521';
my $uid = 'ARCMA';
my $pw = 'arcma';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw, { RaiseError => 1, AutoCommit => 0 }) or die "Unable to connect: $DBI::errstr";

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh, ">", "bypass2.csv" or die "bypass2.csv: $!";

$test = qq{ 
SELECT STORE, DAY, TRAN_SEQ_NO, REV_NO, STORE_DAY_SEQ_NO, TRAN_DATETIME, REGISTER, TRAN_NO, CASHIER, SALESPERSON, TRAN_TYPE, SUB_TRAN_TYPE, ORIG_TRAN_NO, ORIG_TRAN_TYPE, ORIG_REG_NO, REF_NO1, REF_NO2, REF_NO3, REF_NO4, REASON_CODE, VENDOR_NO, VENDOR_INVC_NO, PAYMENT_REF_NO, PROOF_OF_DELIVERY_NO, STATUS, A.VALUE, POS_TRAN_IND, to_char(A.UPDATE_DATETIME,'YYYY-MM-DD HH24:MI:SS') UPDATE_DATETIME, UPDATE_ID, ERROR_IND, BANNER_NO, CUST_ORDER_NO, CUST_ORDER_DATE, ROUNDED_AMT, ROUNDED_OFF_AMT, CREDIT_PROMOTION_ID, REF_NO25, REF_NO26, REF_NO27
FROM SA_TRAN_HEAD@RMS A 
JOIN ADMIN_ETL_SUMMARY B
ON  to_char(A.UPDATE_DATETIME,'YYYY-MM-DD HH24:MI:SS') BETWEEN to_char(TO_date(b.VALUE,'YYYY-MM-DD HH24:MI:SS')+1,'YYYY-MM-DD HH24:MI:SS')
  and TO_char(sysdate,'YYYY-MM-DD HH24:MI:SS')
};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_uc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "bypass2.csv: $!";
	
	$dbh->disconnect; 
}

sub insert_data {

my $truncate = $dbh->prepare( qq{ 
DELETE FROM METRO_IT_SALES_DEPT_LATEPOSTED WHERE TRUNC(UPDATE_DATE) < TRUNC(SYSDATE-5)
});
$truncate->execute();

$truncate->finish();

print "Deleting records from more than 5 days ago table METRO_IT_SALES_DEPT_LATEPOSTED... \nPreparing to Insert... \n";

my $sth_insert = $dbh->prepare( q{
INSERT INTO METRO_IT_SALES_DEPT_LATEPOSTED (STORE, STORE_NAME, MERCH_GROUP_CODE_REV, GROUP_NO, VALUE, UPDATE_DATE)
VALUES ( ?, ?, ?, ?, ?, SYSDATE ) 
});
  
open FH1, '<late_posted_'.$day . '_' . $month_to_text . '_' .$year . '.csv' or die 'Unable to open late_posted_'.$day . '_' . $month_to_text . '_' .$year . '.csv: $!';

<FH1> for 1 .. 1;

while (<FH1>) {
	chomp;
    my ( $store, $store_name, $merch_group_code_rev, $group_no, $value ) = split (/,/);
	
	$sth_insert->execute( $store, $store_name, $merch_group_code_rev, $group_no, $value );	
}
close FH1;

$dbh->commit;
$sth_insert->finish();
$dbh->disconnect;

print "Done with Insert...\nCommited...\n";

}


sub mail {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = 'kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, cham.burgos@metrogaisano.com, marlou.escanilla@metrogaisano.com, joey.labrador@metrogaisano.com, edsel.gayo@metrogaisano.com, Dennis.Cuizon@metrogaisano.com, jeff.bubuli@metrogaisano.com, rex.cabanilla@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'ARC ETL VALIDATION - LATE POSTED ' . $as_of;

$msgbody_file = 'message_BI_LATE.txt';

$attachment_file = 'late_posted_'.$day . '_' . $month_to_text . '_' .$year . '.csv';

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





