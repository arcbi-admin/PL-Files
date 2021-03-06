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
SELECT DIVISION,DEPARTMENT, ACTUALC, BUDGETC, ACHC, ACTUALA, BUDGETA, ACHA FROM (SELECT NULL AS DIVISION, NULL AS DEPARTMENT,'COMP' AS ACTUALC, 'STORES' AS BUDGETC, NULL AS ACHC, 'ALL' AS ACTUALA, 'STORES' AS BUDGETA, NULL AS ACHA FROM DUAL
UNION ALL
SELECT 'DIVISION' AS DIVISION,'DEPARTMENT' AS DEPARTMENT, 'ACTUAL' AS ACTUALC, 'BUDGET' AS BUDGETC, 'ACH' AS ACHC, 'ACTUAL' AS ACTUALA, 'BUDGET' AS BUDGETA, 'ACH' AS ACHA FROM DUAL
UNION ALL
SELECT DIVISION,DEPARTMENT, ACTUAL_COMP ACTUALC, BUDGET_COMP BUDGETC, ACH_COMP ACHC, ACTUAL_ALL ACTUALA, BUDGET_ALL BUDGETA, ACH_ALL ACHA
FROM(SELECT DECODE(GROUPING(MERCH_GROUP_DESC)
             , 0, MERCH_GROUP_DESC
             , 1, 'TOTAL' 
             ) MERCH_GROUP_DESC
, DECODE(GROUPING(DIV_NAME) 
        , 0,  DIV_NAME
        , 1, 'TOTAL' || ' ' || MERCH_GROUP_DESC
        ) DIVISION
, GROUP_NAME AS  DEPARTMENT        
, TO_CHAR(ROUND((SUM(SALE_AMOUNT_COMP)/1000),0),'9G999G999G999') ACTUAL_COMP
, TO_CHAR(SUM(BUDGET_COMP),'9G999G999G999') BUDGET_COMP
, CASE WHEN SUM(BUDGET_COMP) = 0 THEN NULL ELSE TO_CHAR(ROUND(((ROUND((SUM(SALE_AMOUNT_COMP)/1000),2)*100)/(SUM(BUDGET_COMP))),1),'9G999G999G999D9') || '%' END AS ACH_COMP
, TO_CHAR(ROUND((SUM(SALE_AMOUNT_ALL)/1000),0),'9G999G999G999') ACTUAL_ALL
, TO_CHAR(SUM(BUDGET_ALL),'9G999G999G999') BUDGET_ALL
, CASE WHEN SUM(BUDGET_ALL) = 0 THEN NULL ELSE TO_CHAR(ROUND(((ROUND((SUM(SALE_AMOUNT_ALL)/1000),2)*100)/(SUM(BUDGET_ALL))),1),'9G999G999G999D9') || '%' END AS ACH_ALL
FROM
	(SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME,
  SUM(H.MO_SLS_TOT) SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, 0 AS BUDGET_ALL 
	FROM MG_HOURLY_SALES H 
	  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
	  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
	  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
	WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND D.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500') AND TO_NUMBER(H.ID_STR_RT) IN ('6013') AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, SUM(H.MO_SLS_TOT) SALE_AMOUNT_ALL, 0 AS BUDGET_ALL 
	FROM MG_HOURLY_SALES H 
	  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
	  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
	  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
	WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND D.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500') AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME, 0 AS SALE_AMOUNT_COMP, SUM(Q.BUDGET) AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, 0 AS BUDGET_ALL
	FROM MG_BUDGET_DEPT_BI Q 
    LEFT JOIN GROUPS ON Q.DEPT=GROUPS.GROUP_NO
    LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
	WHERE B_DATE = TO_CHAR(SYSDATE, 'DD-MON-YY') AND Q.TYPE = 'C' AND Q.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, SUM(Q.BUDGET) AS BUDGET_ALL
	FROM MG_BUDGET_DEPT_BI Q 
    LEFT JOIN GROUPS ON Q.DEPT=GROUPS.GROUP_NO
    LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
	WHERE B_DATE = TO_CHAR(SYSDATE, 'DD-MON-YY') AND Q.TYPE = 'A' AND Q.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME)
GROUP BY ROLLUP(MERCH_GROUP_DESC, DIV_NAME,GROUP_NAME) 
ORDER BY MERCH_GROUP_DESC, DIV_NAME,GROUP_NAME))

});

$sth->execute() or die "Failed to execute query - " . $dbh->errstr;

my $table = HTML::Table::FromDatabase->new( -sth => $sth, 
                            -border=>0,
                            -width=>'0%',
                            -spacing=>3,
                            -padding=>1,
							);

$table->setSectionId(thead, 0, 'thead' );

$table->delSectionRow(thead, 0, 0);

$table->setRowBGColor(1, '#87CEEB');
$table->setRowBGColor(2, '#87CEEB');
$table->setRowBGColor(3, '#E0FFFF');
$table->setRowBGColor(4, '#E0FFFF');
$table->setRowBGColor(5, '#E0FFFF');
$table->setRowBGColor(6, '#E0FFFF');
$table->setRowBGColor(7, '#E0FFFF');
$table->setRowBGColor(8, '#E0FFFF');
$table->setRowBGColor(9, '#CCEEFF');
$table->setRowBGColor(10, '#E0FFFF');
$table->setRowBGColor(11, '#E0FFFF');
$table->setRowBGColor(12, '#E0FFFF');		
$table->setRowBGColor(13, '#CCEEFF');		
$table->setRowBGColor(14, '#E0FFFF');
$table->setRowBGColor(15, '#E0FFFF');
$table->setRowBGColor(16, '#E0FFFF');
$table->setRowBGColor(17, '#E0FFFF');	
$table->setRowBGColor(18, '#E0FFFF');	
$table->setRowBGColor(19, '#CCEEFF');	
$table->setRowBGColor(20, '#E0FFFF');	
$table->setRowBGColor(21, '#E0FFFF');	
$table->setRowBGColor(22, '#E0FFFF');	
$table->setRowBGColor(23, '#E0FFFF');	
$table->setRowBGColor(24, '#E0FFFF');	
$table->setRowBGColor(25, '#E0FFFF');	
$table->setRowBGColor(26, '#E0FFFF');	
$table->setRowBGColor(27, '#CCEEFF');	
$table->setRowBGColor(28, '#E0FFFF');	
$table->setRowBGColor(29, '#E0FFFF');	
$table->setRowBGColor(30, '#E0FFFF');	
$table->setRowBGColor(31, '#CCEEFF');	
$table->setRowBGColor(32, '#E0FFFF');	
$table->setRowBGColor(33, '#E0FFFF');	
$table->setRowBGColor(34, '#E0FFFF');	
$table->setRowBGColor(35, '#CCEEFF');	
$table->setRowBGColor(36, '#E0FFFF');	
$table->setRowBGColor(37, '#E0FFFF');	
$table->setRowBGColor(38, '#E0FFFF');	
$table->setRowBGColor(39, '#E0FFFF');
$table->setRowBGColor(40, '#E0FFFF');
$table->setRowBGColor(41, '#E0FFFF');	
$table->setRowBGColor(42, '#CCEEFF');	
$table->setRowBGColor(43, '#E0FFFF');	
$table->setRowBGColor(44, '#E0FFFF');	
$table->setRowBGColor(45, '#E0FFFF');	
$table->setRowBGColor(46, '#E0FFFF');	
$table->setRowBGColor(47, '#E0FFFF');	
$table->setRowBGColor(48, '#E0FFFF');	
$table->setRowBGColor(49, '#E0FFFF');	
$table->setRowBGColor(50, '#E0FFFF');	
$table->setRowBGColor(51, '#CCEEFF');	
$table->setRowBGColor(52, '#87CEEB');	
$table->setRowBGColor(53, '#E0FFFF');	
$table->setRowBGColor(54, '#E0FFFF');	
$table->setRowBGColor(55, '#CCEEFF');	
$table->setRowBGColor(56, '#E0FFFF');	
$table->setRowBGColor(57, '#E0FFFF');	
$table->setRowBGColor(58, '#E0FFFF');	
$table->setRowBGColor(59, '#E0FFFF');
$table->setRowBGColor(60, '#E0FFFF');
$table->setRowBGColor(61, '#E0FFFF');	
$table->setRowBGColor(62, '#CCEEFF');	
$table->setRowBGColor(63, '#E0FFFF');	
$table->setRowBGColor(64, '#E0FFFF');	
$table->setRowBGColor(65, '#E0FFFF');	
$table->setRowBGColor(66, '#CCEEFF');	
$table->setRowBGColor(67, '#E0FFFF');	
$table->setRowBGColor(68, '#E0FFFF');	
$table->setRowBGColor(69, '#E0FFFF');
$table->setRowBGColor(70, '#E0FFFF');
$table->setRowBGColor(71, '#E0FFFF');	
$table->setRowBGColor(72, '#CCEEFF');
$table->setRowBGColor(73, '#E0FFFF');

$table->setRowBGColor(74, '#CCEEFF');
$table->setRowBGColor(75, '#87CEEB');
$table->setRowBGColor(76, '#87CEEB');



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


$table->setCellAlign(3, 2, 'left');
$table->setCellAlign(4, 2, 'left');
$table->setCellAlign(5, 2, 'left');
$table->setCellAlign(6, 2, 'left');
$table->setCellAlign(7, 2, 'left');
$table->setCellAlign(8, 2, 'left');
$table->setCellAlign(9, 2, 'left');
$table->setCellAlign(10, 2, 'left');
$table->setCellAlign(11, 2, 'left');
$table->setCellAlign(12, 2, 'left');
$table->setCellAlign(13, 2, 'left');
$table->setCellAlign(14, 2, 'left');
$table->setCellAlign(15, 2, 'left');
$table->setCellAlign(16, 2, 'left');
$table->setCellAlign(17, 2, 'left');
$table->setCellAlign(18, 2, 'left');
$table->setCellAlign(19, 2, 'left');
$table->setCellAlign(20, 2, 'left');
$table->setCellAlign(21, 2, 'left');
$table->setCellAlign(22, 2, 'left');
$table->setCellAlign(23, 2, 'left');
$table->setCellAlign(24, 2, 'left');
$table->setCellAlign(25, 2, 'left');
$table->setCellAlign(26, 2, 'left');
$table->setCellAlign(27, 2, 'left');
$table->setCellAlign(28, 2, 'left');
$table->setCellAlign(29, 2, 'left');
$table->setCellAlign(30, 2, 'left');
$table->setCellAlign(31, 2, 'left');
$table->setCellAlign(32, 2, 'left');
$table->setCellAlign(33, 2, 'left');
$table->setCellAlign(34, 2, 'left');
$table->setCellAlign(35, 2, 'left');
$table->setCellAlign(36, 2, 'left');
$table->setCellAlign(37, 2, 'left');
$table->setCellAlign(38, 2, 'left');
$table->setCellAlign(39, 2, 'left');
$table->setCellAlign(40, 2, 'left');
$table->setCellAlign(41, 2, 'left');
$table->setCellAlign(42, 2, 'left');
$table->setCellAlign(43, 2, 'left');
$table->setCellAlign(44, 2, 'left');
$table->setCellAlign(45, 2, 'left');
$table->setCellAlign(46, 2, 'left');
$table->setCellAlign(47, 2, 'left');
$table->setCellAlign(48, 2, 'left');
$table->setCellAlign(49, 2, 'left');
$table->setCellAlign(50, 2, 'left');
$table->setCellAlign(51, 2, 'left');
$table->setCellAlign(52, 2, 'left');
$table->setCellAlign(53, 2, 'left');
$table->setCellAlign(54, 2, 'left');
$table->setCellAlign(55, 2, 'left');
$table->setCellAlign(56, 2, 'left');
$table->setCellAlign(57, 2, 'left');
$table->setCellAlign(58, 2, 'left');
$table->setCellAlign(59, 2, 'left');
$table->setCellAlign(60, 2, 'left');
$table->setCellAlign(61, 2, 'left');
$table->setCellAlign(62, 2, 'left');
$table->setCellAlign(63, 2, 'left');
$table->setCellAlign(64, 2, 'left');
$table->setCellAlign(65, 2, 'left');
$table->setCellAlign(66, 2, 'left');
$table->setCellAlign(67, 2, 'left');
$table->setCellAlign(68, 2, 'left');
$table->setCellAlign(69, 2, 'left');
$table->setCellAlign(70, 2, 'left');
$table->setCellAlign(71, 2, 'left');
$table->setCellAlign(72, 2, 'left');
$table->setCellAlign(73, 2, 'left');




my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject ) = @ARGV;

#$to = ' chit.lazaro@metrogaisano.com, fili.mercado@metrogaisano.com, karan.malani@metrogaisano.com, lia.chipeco@metrogaisano.com, marlita.portes@metrogaisano.com, jordan.mok@metrogaisano.com, jennifer.yu@metrogaisano.com, april.agapito@metrogaisano.com, edna.prieto@metrogaisano.com, tessie.baldezamo@metrogaisano.com, chedie.lim@metrogaisano.com,jennifer.nardo@metrogaisano.com, liberato.rodriguez@metrogaisano.com, cj.jesena@metrogaisano.com, luz.bitang@metrogaisano.com, emily.silverio@metrogaisano.com, glenda.navares@metrogaisano.com, julie.montano@metrogaisano.com, may.sasedor@metrogaisano.com, roy.igot@metrogaisano.com, limuel.ulanday@metrogaisano.com, delia.jakosalem@metrogaisano.com ';

$cc = 'rex.cabanilla@metrogaisano.com, annalyn.conde@metrogaisano.com ';

$bcc = 'lea.gonzaga@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com';
#$bcc = ' lea.gonzaga@metrogaisano.com';
#$bcc = ' lea.gonzaga@metrogaisano.com';


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
<b><font size="4">Hourly Sales Report</font></b>
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

