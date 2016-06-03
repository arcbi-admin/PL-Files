use DBI;
use DBD::Oracle qw(:ora_types);
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use Text::CSV_XS;
use DBConnector;
use Win32::Job;
use Getopt::Long;
use IO::File;
use MIME::QuotedPrint;
use MIME::Base64;
use Mail::Sendmail;
use HTML::Entities;
use HTML::Table::FromDatabase;
use CGI;
use HTML::Template;

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

$test_update = qq{ SELECT CUR_TIME, UPDATE_TIME, CASE WHEN CUR_TIME <> UPDATE_TIME THEN 0 ELSE 1 END AS STATUS 
					FROM 
					(SELECT TO_CHAR(SYSDATE, 'HH24') CUR_TIME FROM DUAL)A,
					(SELECT CASE WHEN NEW_TIME IS NULL THEN '0' ELSE TO_CHAR(NEW_TIME) END AS UPDATE_TIME 
					  FROM ( SELECT MAX(TO_NUMBER(TS_RTN_HR)) NEW_TIME
										FROM MG_HOURLY_SALES WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND TO_NUMBER(ID_STR_RT) = '4004'))B };

$test_update = $dbh->prepare($test_update);
$test_update->execute();

while ( my $x =  $test_update->fetchrow_hashref()){
	$test = $x->{STATUS};
} 

if ($test eq 1){				
	$test_query = qq{ SELECT CASE WHEN NEW_TIME > 12 THEN  TO_CHAR(NEW_TIME -12) || ':00 PM' ELSE TO_CHAR(NEW_TIME) || ':00 AM' END AS UPDATE_TIME FROM (
						SELECT MAX(TO_NUMBER(TS_RTN_HR)) NEW_TIME
						FROM MG_HOURLY_SALES WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND TO_NUMBER(ID_STR_RT) = '4004') };

	$tst_query = $dbh->prepare($test_query);
	$tst_query->execute();

	while ( my $x =  $tst_query->fetchrow_hashref()){
		$update_time = $x->{UPDATE_TIME};
	} 

	&mailer;
	&mailer_external;

	$tst_query->finish();
	$test_update->finish();
	$dbh->disconnect;
	
}

elsif ($test eq 0){
	$test_update->finish();
	$dbh->disconnect;
	print "Not updated...\nExiting...\n";
	exit;
}

sub mailer {

 
my $sth = $dbh->prepare(qq{
SELECT DIVISION, ACTUALC, BUDGETC, ACHC 
FROM (
	SELECT 'DIVISION' AS DIVISION, 'ACTUAL' AS ACTUALC, 'BUDGET' AS BUDGETC, 'ACH' AS ACHC FROM DUAL
	UNION ALL
	SELECT DIVISION, ACTUAL_COMP ACTUALC, BUDGET_COMP BUDGETC, ACH_COMP ACHC
	FROM
	(SELECT 
	DECODE(GROUPING(MERCH_GROUP_DESC)
				 , 0, MERCH_GROUP_DESC
				 , 1, 'TOTAL' 
				 ) MERCH_GROUP_DESC
	, DECODE(GROUPING(DIV_NAME) 
			, 0,  DIV_NAME
			, 1, 'TOTAL' || ' ' || MERCH_GROUP_DESC
			) DIVISION
	, TO_CHAR(ROUND((SUM(SALE_AMOUNT_COMP)/1000),0),'9G999G999G999') ACTUAL_COMP
	, TO_CHAR(SUM(BUDGET_COMP),'9G999G999G999') BUDGET_COMP
	, CASE WHEN SUM(BUDGET_COMP) = 0 THEN NULL ELSE TO_CHAR(ROUND(((ROUND((SUM(SALE_AMOUNT_COMP)/1000),2)*100)/(SUM(BUDGET_COMP))),1),'9G999G999G999D9') || '%' END AS ACH_COMP
	FROM
		(SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, SUM(H.MO_SLS_TOT) SALE_AMOUNT_COMP, 0 AS BUDGET_COMP
		FROM MG_HOURLY_SALES H 
		  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
		  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
		  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
		  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
		WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND D.DIVISION IN ('1500','6000') AND TO_NUMBER(H.ID_STR_RT) = '4004' AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
		GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
	UNION ALL
		SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, SUM(Q.BUDGET) AS BUDGET_COMP
		FROM MG_Q4_BUDGET_BI Q 
		LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
		  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
		WHERE TO_DATE(Q.B_DATE,'MM/DD/YY') = TO_CHAR(SYSDATE, 'DD-MON-YY') AND Q.TYPE = '4004' AND Q.DIVISION IN ('1500','6000') 
		GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME)
	GROUP BY ROLLUP(MERCH_GROUP_DESC, DIV_NAME) 
	ORDER BY MERCH_GROUP_DESC, DIV_NAME)
)
});

$sth->execute() or die "Failed to execute query - " . $dbh->errstr;

my $table = HTML::Table::FromDatabase->new( -sth => $sth, 
                            -border=>0,
                            -width=>'0%',
                            -spacing=>3,
                            -padding=>2,
							);

$table->setSectionId(thead, 0, 'thead' );

$table->delSectionRow(thead, 0, 0);

$table->setRowBGColor(1, '#87CEEB');
$table->setRowBGColor(2, '#C0C0C0');
$table->setRowBGColor(3, '#C0C0C0');
$table->setRowBGColor(4, '#C0C0C0');
$table->setRowBGColor(5, '#C0C0C0');
$table->setRowBGColor(6, '#C0C0C0');
$table->setRowBGColor(7, '#C0C0C0');
$table->setRowBGColor(8, '#C0C0C0');
$table->setRowBGColor(9, '#C0C0C0');
$table->setRowBGColor(10, '#87CEEB');
$table->setRowBGColor(11, '#C0C0C0');		
$table->setRowBGColor(12, '#C0C0C0');		
$table->setRowBGColor(13, '#C0C0C0');
$table->setRowBGColor(14, '#C0C0C0');
$table->setRowBGColor(15, '#C0C0C0');	
$table->setRowBGColor(16, '#87CEEB');
$table->setRowBGColor(17, '#87CEEB');

$table->setColAlign(2, 'right');
$table->setColAlign(3, 'right');
$table->setColAlign(4, 'right');

$table->setCellAlign(1, 1, 'center');
$table->setCellAlign(1, 2, 'center');
$table->setCellAlign(1, 3, 'center');
$table->setCellAlign(1, 4, 'center');


my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject ) = @ARGV;

$to = ' emily.silverio@metrogaisano.com, limuel.ulanday@metrogaisano.com, maricel.tamala@metrogaisano.com, fili.mercado@metrogaisano.com, chit.lazaro@metrogaisano.com, karan.malani@metrogaisano.com, lia.chipeco@metrogaisano.com, luz.bitang@metrogaisano.com ';

$cc = ' arthur.emmanuel@metrogaisano.com, rex.cabanilla@metrogaisano.com, annalyn.conde@metrogaisano.com ';

$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, cham.burgos@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';

$subject = 'Hourly Sales Performance as of ' . $update_time;

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
Content-Type: text/html; charset="iso-8859-1"

<html>
<b><font size="4">Hourly Sales Report </font></b><br>
<b><font size="3">4004 - METRO NEWPORT PASAY</font></b><br>
As of &nbsp;$update_time<br><br>

$table1
$table

<p><i><font size="2">in 000s</font></i>.</p><br>

Regards, <br>
<a href= "mailto:arcbi.support@metrogaisano.com">ARC BI Support</a>
</html>

$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

$sth->finish();

}

sub mailer_external {
 
my $sth = $dbh->prepare(qq{
SELECT DIVISION, ACTUALC, BUDGETC, ACHC 
FROM (
	SELECT 'DIVISION' AS DIVISION, 'ACTUAL' AS ACTUALC, 'BUDGET' AS BUDGETC, 'ACH' AS ACHC FROM DUAL
	UNION ALL
	SELECT DIVISION, ACTUAL_COMP ACTUALC, BUDGET_COMP BUDGETC, ACH_COMP ACHC
	FROM
	(SELECT 
	DECODE(GROUPING(MERCH_GROUP_DESC)
				 , 0, MERCH_GROUP_DESC
				 , 1, 'TOTAL' 
				 ) MERCH_GROUP_DESC
	, DECODE(GROUPING(DIV_NAME) 
			, 0,  DIV_NAME
			, 1, 'TOTAL' || ' ' || MERCH_GROUP_DESC
			) DIVISION
	, TO_CHAR(ROUND((SUM(SALE_AMOUNT_COMP)/1000),0),'9G999G999G999') ACTUAL_COMP
	, TO_CHAR(SUM(BUDGET_COMP),'9G999G999G999') BUDGET_COMP
	, CASE WHEN SUM(BUDGET_COMP) = 0 THEN NULL ELSE TO_CHAR(ROUND(((ROUND((SUM(SALE_AMOUNT_COMP)/1000),2)*100)/(SUM(BUDGET_COMP))),1),'9G999G999G999D9') || '%' END AS ACH_COMP
	FROM
		(SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, SUM(H.MO_SLS_TOT) SALE_AMOUNT_COMP, 0 AS BUDGET_COMP
		FROM MG_HOURLY_SALES H 
		  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
		  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
		  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
		  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
		WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND D.DIVISION IN ('1500','6000') AND TO_NUMBER(H.ID_STR_RT) = '4004' AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
		GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
	UNION ALL
		SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, SUM(Q.BUDGET) AS BUDGET_COMP
		FROM MG_Q4_BUDGET_BI Q 
		LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
		  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
		WHERE TO_DATE(Q.B_DATE,'MM/DD/YY') = TO_CHAR(SYSDATE, 'DD-MON-YY') AND Q.TYPE = '4004' AND Q.DIVISION IN ('1500','6000') 
		GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME)
	GROUP BY ROLLUP(MERCH_GROUP_DESC, DIV_NAME) 
	ORDER BY MERCH_GROUP_DESC, DIV_NAME)
)
});

$sth->execute() or die "Failed to execute query - " . $dbh->errstr;

my $table = HTML::Table::FromDatabase->new( -sth => $sth, 
                            -border=>0,
                            -width=>'0%',
                            -spacing=>3,
                            -padding=>2,
							);

$table->setSectionId(thead, 0, 'thead' );

$table->delSectionRow(thead, 0, 0);

$table->setRowBGColor(1, '#87CEEB');
$table->setRowBGColor(2, '#C0C0C0');
$table->setRowBGColor(3, '#C0C0C0');
$table->setRowBGColor(4, '#C0C0C0');
$table->setRowBGColor(5, '#C0C0C0');
$table->setRowBGColor(6, '#C0C0C0');
$table->setRowBGColor(7, '#C0C0C0');
$table->setRowBGColor(8, '#C0C0C0');
$table->setRowBGColor(9, '#C0C0C0');
$table->setRowBGColor(10, '#87CEEB');
$table->setRowBGColor(11, '#C0C0C0');		
$table->setRowBGColor(12, '#C0C0C0');		
$table->setRowBGColor(13, '#C0C0C0');
$table->setRowBGColor(14, '#C0C0C0');
$table->setRowBGColor(15, '#C0C0C0');	
$table->setRowBGColor(16, '#87CEEB');
$table->setRowBGColor(17, '#87CEEB');

$table->setColAlign(2, 'right');
$table->setColAlign(3, 'right');
$table->setColAlign(4, 'right');

$table->setCellAlign(1, 1, 'center');
$table->setCellAlign(1, 2, 'center');
$table->setCellAlign(1, 3, 'center');
$table->setCellAlign(1, 4, 'center');

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject ) = @ARGV;

$to = ' artemm12@aol.com, frankgaisano@gmail.com ';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';

$subject = 'Hourly Sales Performance as of ' . $update_time;

my %mail = (
    To   => $to,
	From  => $from,
    Subject => $subject,
	'content-type' => "multipart/alternative; boundary=\"$boundary\""
);

$mail{smtp} = '10.190.1.54';
$mail{Cc} = $cc if $cc;
$mail{Bcc} = $bcc if $bcc;

my $boundary = "====" . time . "====";

$mail{'content-type'} = qq(multipart/mixed; boundary="$boundary");

$boundary = '--'.$boundary;

$mail{body} = 
<<END_OF_BODY;
$boundary
Content-Type: text/html; charset="iso-8859-1"

<html>
<b><font size="4">Hourly Sales Report</font></b><br>
<b><font size="3">4004 - METRO NEWPORT PASAY</font></b><br>
As of &nbsp;$update_time<br><br>

$table1
$table

<p><i><font size="2">in 000s</font></i>.</p><br>

Regards, <br>
<a href= "mailto:arcbi.support@metrogaisano.com">ARC BI Support</a>
</html>

$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

$sth->finish();

}

sub read_file {

my( $filename, $binmode ) = @_;
my $fh = new IO::File;
$fh->open("<".$filename) or die "Error opening $filename for reading - $!\n";
$fh->binmode if $binmode;
local $/;
<$fh>
	
}

