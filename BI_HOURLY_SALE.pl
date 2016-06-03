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

&mailer;
&mailer_external;

sub mailer {

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

#$test_query = qq{ SELECT CASE WHEN NEW_TIME > 12 THEN  TO_CHAR(NEW_TIME -12) || ':14 PM' ELSE TO_CHAR(NEW_TIME) || ':14 AM' END AS UPDATE_TIME FROM (
#					SELECT MAX(TO_NUMBER(TS_RTN_HR)) NEW_TIME
#					FROM MG_HOURLY_SALES WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD'))) };

$test_query = qq{ SELECT CASE WHEN NEW_TIME > 12 THEN  TO_CHAR(NEW_TIME -12) || ':14 PM' ELSE TO_CHAR(NEW_TIME) || ':14 AM' END AS UPDATE_TIME FROM (
					SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL) };

$tst_query = $dbh->prepare($test_query);
$tst_query->execute();

while ( my $x =  $tst_query->fetchrow_hashref()){
	$update_time = $x->{UPDATE_TIME};
} 
 
my $sth = $dbh->prepare(qq{
SELECT DIVISION, ACTUALC, BUDGETC, ACHC, ACTUALA, BUDGETA, ACHA FROM (
SELECT NULL AS DIVISION, 'COMP' AS ACTUALC, 'STORES' AS BUDGETC, NULL AS ACHC, 'ALL' AS ACTUALA, 'STORES' AS BUDGETA, NULL AS ACHA FROM DUAL
UNION ALL
SELECT 'DIVISION' AS DIVISION, 'ACTUAL' AS ACTUALC, 'BUDGET' AS BUDGETC, 'ACH' AS ACHC, 'ACTUAL' AS ACTUALA, 'BUDGET' AS BUDGETA, 'ACH' AS ACHA FROM DUAL
UNION ALL
SELECT DIVISION, ACTUAL_COMP ACTUALC, BUDGET_COMP BUDGETC, ACH_COMP ACHC, ACTUAL_ALL ACTUALA, BUDGET_ALL BUDGETA, ACH_ALL ACHA
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
, TO_CHAR(ROUND((SUM(SALE_AMOUNT_ALL)/1000),0),'9G999G999G999') ACTUAL_ALL
, TO_CHAR(SUM(BUDGET_ALL),'9G999G999G999') BUDGET_ALL
, CASE WHEN SUM(BUDGET_ALL) = 0 THEN NULL ELSE TO_CHAR(ROUND(((ROUND((SUM(SALE_AMOUNT_ALL)/1000),2)*100)/(SUM(BUDGET_ALL))),1),'9G999G999G999D9') || '%' END AS ACH_ALL
FROM
	(SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, SUM(H.MO_SLS_TOT) SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, 0 AS BUDGET_ALL 
	FROM MG_HOURLY_SALES H 
	  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
	  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
	  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
	WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND D.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500') AND TO_NUMBER(H.ID_STR_RT) IN ('2009','2012','7176','7003','2006','7004','2007','6008','7005','2010','7300','7009','5006','5005','5004','5003','5002','5001','7173','3004','3003','3001','4003','7008','7007','7006','3006','3007','2003','7000','4002','3005','3002','2002','2001','2011','2008','7001','2004','7002','2005', '3012') AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, SUM(H.MO_SLS_TOT) SALE_AMOUNT_ALL, 0 AS BUDGET_ALL 
	FROM MG_HOURLY_SALES H 
	  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
	  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
	  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
	WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND D.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500') AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, SUM(Q.BUDGET) AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, 0 AS BUDGET_ALL
	FROM MG_Q4_BUDGET_BI Q 
    LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
	WHERE TO_DATE(Q.B_DATE,'MM/DD/YY') = TO_CHAR(SYSDATE, 'DD-MON-YY') AND Q.TYPE = 'C' AND Q.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, SUM(Q.BUDGET) AS BUDGET_ALL
	FROM MG_Q4_BUDGET_BI Q 
    LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
	WHERE TO_DATE(Q.B_DATE,'MM/DD/YY') = TO_CHAR(SYSDATE, 'DD-MON-YY') AND Q.TYPE = 'A' AND Q.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME)
GROUP BY ROLLUP(MERCH_GROUP_DESC, DIV_NAME) 
ORDER BY MERCH_GROUP_DESC, DIV_NAME
))
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
$table->setRowBGColor(2, '#87CEEB');
$table->setRowBGColor(3, '#C0C0C0');
$table->setRowBGColor(4, '#C0C0C0');
$table->setRowBGColor(5, '#C0C0C0');
$table->setRowBGColor(6, '#C0C0C0');
$table->setRowBGColor(7, '#C0C0C0');
$table->setRowBGColor(8, '#C0C0C0');
$table->setRowBGColor(9, '#C0C0C0');
$table->setRowBGColor(10, '#C0C0C0');
$table->setRowBGColor(11, '#87CEEB');
$table->setRowBGColor(12, '#C0C0C0');		
$table->setRowBGColor(13, '#C0C0C0');		
$table->setRowBGColor(14, '#C0C0C0');
$table->setRowBGColor(15, '#C0C0C0');
$table->setRowBGColor(16, '#C0C0C0');	
$table->setRowBGColor(17, '#87CEEB');
$table->setRowBGColor(18, '#87CEEB');

$table->setColAlign(2, 'right');
$table->setColAlign(3, 'right');
$table->setColAlign(4, 'right');
$table->setColAlign(5, 'right');
$table->setColAlign(6, 'right');
$table->setColAlign(7, 'right');

$table->setCellAlign(1, 2, 'right');
$table->setCellAlign(1, 3, 'left');
$table->setCellAlign(1, 5, 'right');
$table->setCellAlign(1, 6, 'left');

$table->setCellAlign(2, 1, 'center');
$table->setCellAlign(2, 2, 'center');
$table->setCellAlign(2, 3, 'center');
$table->setCellAlign(2, 4, 'center');
$table->setCellAlign(2, 5, 'center');
$table->setCellAlign(2, 6, 'center');
$table->setCellAlign(2, 7, 'center');

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject ) = @ARGV;

$to = ' chit.lazaro@metrogaisano.com, fili.mercado@metrogaisano.com, karan.malani@metrogaisano.com, lia.chipeco@metrogaisano.com, marlita.portes@metrogaisano.com, jordan.mok@metrogaisano.com, patricia.canton@metrogaisano.com, jennifer.yu@metrogaisano.com, april.agapito@metrogaisano.com, edna.prieto@metrogaisano.com, tessie.baldezamo@metrogaisano.com, chedie.lim@metrogaisano.com,jennifer.nardo@metrogaisano.com, liberato.rodriguez@metrogaisano.com, cj.jesena@metrogaisano.com, luz.bitang@metrogaisano.com, emily.silverio@metrogaisano.com, glenda.navares@metrogaisano.com, julie.montano@metrogaisano.com, may.sasedor@metrogaisano.com, alain.reyes@metrogaisano.com, roy.igot@metrogaisano.com, limuel.ulanday@metrogaisano.com, jacqueline.cano@metrogaisano.com, joefrey.camu@metrogaisano.com, dinah.ramirez@metrogaisano.com, delia.jakosalem@metrogaisano.com ';

$cc = ' arthur.emmanuel@metrogaisano.com, frank.gaisano@metrogaisano.com, eric.redona@metrogaisano.com, rex.cabanilla@metrogaisano.com, annalyn.conde@metrogaisano.com ';

$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com';

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
<b><font size="4">Hourly Sales Report</font></b><br>
<i><font size="2">Excluding 6 TG Stores, Fresh 'n Easy Mactan and Taguig, Carcar and LapuLapu Hypermarkets</font></i><br>
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
$dbh->disconnect;

}

sub mailer_external {

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

# $test_query = qq{ SELECT CASE WHEN NEW_TIME > 12 THEN  TO_CHAR(NEW_TIME -12) || ':14 PM' ELSE TO_CHAR(NEW_TIME) || ':14 AM' END AS UPDATE_TIME FROM (
					# SELECT MAX(TO_NUMBER(TS_RTN_HR)) NEW_TIME
					# FROM MG_HOURLY_SALES WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD'))) };
					
$test_query = qq{ SELECT CASE WHEN NEW_TIME > 12 THEN  TO_CHAR(NEW_TIME -12) || ':14 PM' ELSE TO_CHAR(NEW_TIME) || ':14 AM' END AS UPDATE_TIME FROM (
					SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL) };

$tst_query = $dbh->prepare($test_query);
$tst_query->execute();

while ( my $x =  $tst_query->fetchrow_hashref()){
	$update_time = $x->{UPDATE_TIME};
} 
 
my $sth = $dbh->prepare(qq{
SELECT DIVISION, ACTUALC, BUDGETC, ACHC, ACTUALA, BUDGETA, ACHA FROM (
SELECT NULL AS DIVISION, 'COMP' AS ACTUALC, 'STORES' AS BUDGETC, NULL AS ACHC, 'ALL' AS ACTUALA, 'STORES' AS BUDGETA, NULL AS ACHA FROM DUAL
UNION ALL
SELECT 'DIVISION' AS DIVISION, 'ACTUAL' AS ACTUALC, 'BUDGET' AS BUDGETC, 'ACH' AS ACHC, 'ACTUAL' AS ACTUALA, 'BUDGET' AS BUDGETA, 'ACH' AS ACHA FROM DUAL
UNION ALL
SELECT DIVISION, ACTUAL_COMP ACTUALC, BUDGET_COMP BUDGETC, ACH_COMP ACHC, ACTUAL_ALL ACTUALA, BUDGET_ALL BUDGETA, ACH_ALL ACHA
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
, TO_CHAR(ROUND((SUM(SALE_AMOUNT_ALL)/1000),0),'9G999G999G999') ACTUAL_ALL
, TO_CHAR(SUM(BUDGET_ALL),'9G999G999G999') BUDGET_ALL
, CASE WHEN SUM(BUDGET_ALL) = 0 THEN NULL ELSE TO_CHAR(ROUND(((ROUND((SUM(SALE_AMOUNT_ALL)/1000),2)*100)/(SUM(BUDGET_ALL))),1),'9G999G999G999D9') || '%' END AS ACH_ALL
FROM
	(SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, SUM(H.MO_SLS_TOT) SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, 0 AS BUDGET_ALL 
	FROM MG_HOURLY_SALES H 
	  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
	  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
	  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
	WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND D.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500') AND TO_NUMBER(H.ID_STR_RT) IN ('2009','2012','7176','7003','2006','7004','2007','6008','7005','2010','7300','7009','5006','5005','5004','5003','5002','5001','7173','3004','3003','3001','4003','7008','7007','7006','3006','3007','2003','7000','4002','3005','3002','2002','2001','2011','2008','7001','2004','7002','2005', '3012') AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, SUM(H.MO_SLS_TOT) SALE_AMOUNT_ALL, 0 AS BUDGET_ALL 
	FROM MG_HOURLY_SALES H 
	  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
	  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
	  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
	WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND D.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500') AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, SUM(Q.BUDGET) AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, 0 AS BUDGET_ALL
	FROM MG_Q4_BUDGET_BI Q 
    LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
	WHERE TO_DATE(Q.B_DATE,'MM/DD/YY') = TO_CHAR(SYSDATE, 'DD-MON-YY') AND Q.TYPE = 'C' AND Q.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, SUM(Q.BUDGET) AS BUDGET_ALL
	FROM MG_Q4_BUDGET_BI Q 
    LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
	WHERE TO_DATE(Q.B_DATE,'MM/DD/YY') = TO_CHAR(SYSDATE, 'DD-MON-YY') AND Q.TYPE = 'A' AND Q.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME)
GROUP BY ROLLUP(MERCH_GROUP_DESC, DIV_NAME) 
ORDER BY MERCH_GROUP_DESC, DIV_NAME
))
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
$table->setRowBGColor(2, '#87CEEB');
$table->setRowBGColor(3, '#C0C0C0');
$table->setRowBGColor(4, '#C0C0C0');
$table->setRowBGColor(5, '#C0C0C0');
$table->setRowBGColor(6, '#C0C0C0');
$table->setRowBGColor(7, '#C0C0C0');
$table->setRowBGColor(8, '#C0C0C0');
$table->setRowBGColor(9, '#C0C0C0');
$table->setRowBGColor(10, '#C0C0C0');
$table->setRowBGColor(11, '#87CEEB');
$table->setRowBGColor(12, '#C0C0C0');		
$table->setRowBGColor(13, '#C0C0C0');		
$table->setRowBGColor(14, '#C0C0C0');
$table->setRowBGColor(15, '#C0C0C0');
$table->setRowBGColor(16, '#C0C0C0');	
$table->setRowBGColor(17, '#87CEEB');
$table->setRowBGColor(18, '#87CEEB');

$table->setColAlign(2, 'right');
$table->setColAlign(3, 'right');
$table->setColAlign(4, 'right');
$table->setColAlign(5, 'right');
$table->setColAlign(6, 'right');
$table->setColAlign(7, 'right');

$table->setCellAlign(1, 2, 'right');
$table->setCellAlign(1, 3, 'left');
$table->setCellAlign(1, 5, 'right');
$table->setCellAlign(1, 6, 'left');

$table->setCellAlign(2, 1, 'center');
$table->setCellAlign(2, 2, 'center');
$table->setCellAlign(2, 3, 'center');
$table->setCellAlign(2, 4, 'center');
$table->setCellAlign(2, 5, 'center');
$table->setCellAlign(2, 6, 'center');
$table->setCellAlign(2, 7, 'center');

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject ) = @ARGV;

$to = 'artemm12@aol.com, frankgaisano@gmail.com';

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
<i><font size="2">Excluding 6 TG Stores, Fresh 'n Easy Mactan and Taguig, Carcar and LapuLapu Hypermarkets</font></i><br>
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
$dbh->disconnect;

}

sub read_file {

my( $filename, $binmode ) = @_;
my $fh = new IO::File;
$fh->open("<".$filename) or die "Error opening $filename for reading - $!\n";
$fh->binmode if $binmode;
local $/;
<$fh>
	
}

