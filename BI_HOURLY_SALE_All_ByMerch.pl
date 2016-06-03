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

$test_query = qq{ SELECT CASE WHEN NEW_TIME > 12 THEN  TO_CHAR(NEW_TIME -12) || ':15 PM' ELSE TO_CHAR(NEW_TIME) || ':15 AM' END AS UPDATE_TIME FROM (
					SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL) };

$tst_query = $dbh->prepare($test_query);
$tst_query->execute();

while ( my $x =  $tst_query->fetchrow_hashref()){
	$update_time = $x->{UPDATE_TIME};
} 
 
my $sth = $dbh->prepare(qq{
SELECT DIVISION,DEPARTMENT, ACTUALA, BUDGETA, ACHA FROM 
(SELECT 'DIVISION' AS DIVISION,'DEPARTMENT' AS DEPARTMENT, 'ACTUAL' AS ACTUALA, 'BUDGET' AS BUDGETA, 'ACH' AS ACHA FROM DUAL
UNION ALL
SELECT DIVISION,DEPARTMENT,  ACTUAL_ALL ACTUALA, BUDGET_ALL BUDGETA, ACH_ALL ACHA
FROM(SELECT DECODE(GROUPING(MERCH_GROUP_DESC)
             , 0, MERCH_GROUP_DESC
             , 1, 'TOTAL' 
             ) MERCH_GROUP_DESC
, DECODE(GROUPING(DIV_NAME) 
        , 0,  DIV_NAME
        , 1, 'TOTAL' || ' ' || MERCH_GROUP_DESC
        ) DIVISION
, GROUP_NAME AS  DEPARTMENT        
, TO_CHAR(ROUND((SUM(SALE_AMOUNT_ALL)/1000),0),'9G999G999G999') ACTUAL_ALL
, TO_CHAR(SUM(BUDGET_ALL),'9G999G999G999') BUDGET_ALL
, CASE WHEN SUM(BUDGET_ALL) = 0 THEN NULL ELSE TO_CHAR(ROUND(((ROUND((SUM(SALE_AMOUNT_ALL)/1000),2)*100)/(SUM(BUDGET_ALL))),1),'9G999G999G999D9') || '%' END AS ACH_ALL
FROM
	(SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, SUM(H.MO_SLS_TOT) SALE_AMOUNT_ALL, 0 AS BUDGET_ALL 
	FROM MG_HOURLY_SALES H 
	  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
	  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
	  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
	WHERE DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD')) AND D.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	--AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
	--AND H.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME,GROUPS.GROUP_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, 
	0 AS SALE_AMOUNT_ALL, SUM(Q.BUDGET) AS BUDGET_ALL
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

$table->setRowBGColor(1, '#81BEF7');
#$table->setRowBGColor(1, '#B3FFB3');
$table->setRowBGColor(2, '#E6FFE6');
$table->setRowBGColor(2, '#E6FFE6');
$table->setRowBGColor(3, '#E6FFE6');
$table->setRowBGColor(4, '#E6FFE6'); #dri kutob
$table->setRowBGColor(5, '#E6FFE6');
$table->setRowBGColor(6, '#E6FFE6');
$table->setRowBGColor(7, '#E6FFE6');
$table->setRowBGColor(8, '#E6FFE6');
$table->setRowBGColor(9, '#E6FFE6');
$table->setRowBGColor(10, '#E6FFE6');
$table->setRowBGColor(11, '#E6FFE6');
$table->setRowBGColor(12, '#E6FFE6');		
$table->setRowBGColor(13, '#E6FFE6');		
$table->setRowBGColor(14, '#E6FFE6');
$table->setRowBGColor(15, '#E6FFE6');
$table->setRowBGColor(16, '#E6FFE6');
$table->setRowBGColor(17, '#E6FFE6');	
$table->setRowBGColor(18, '#E6FFE6');	
$table->setRowBGColor(19, '#E6FFE6');	
$table->setRowBGColor(20, '#E6FFE6');	
$table->setRowBGColor(21, '#E6FFE6');	
$table->setRowBGColor(22, '#E6FFE6');	
$table->setRowBGColor(23, '#E6FFE6');	
$table->setRowBGColor(24, '#E6FFE6');	
$table->setRowBGColor(25, '#E6FFE6');	
$table->setRowBGColor(26, '#E6FFE6');	
$table->setRowBGColor(27, '#E6FFE6');	
$table->setRowBGColor(28, '#E6FFE6');	
$table->setRowBGColor(29, '#E6FFE6');	
$table->setRowBGColor(30, '#E6FFE6');	
$table->setRowBGColor(31, '#E6FFE6');	
$table->setRowBGColor(32, '#E6FFE6');	
$table->setRowBGColor(33, '#E6FFE6');	
$table->setRowBGColor(34, '#E6FFE6');	
$table->setRowBGColor(35, '#E6FFE6');	
$table->setRowBGColor(36, '#E6FFE6');	
$table->setRowBGColor(37, '#E6FFE6');	
$table->setRowBGColor(38, '#E6FFE6');	
$table->setRowBGColor(39, '#E6FFE6');
$table->setRowBGColor(40, '#E6FFE6');
$table->setRowBGColor(41, '#E6FFE6');	
$table->setRowBGColor(42, '#E6FFE6');	
$table->setRowBGColor(43, '#E6FFE6');	
$table->setRowBGColor(44, '#E6FFE6');	
$table->setRowBGColor(45, '#E6FFE6');	
$table->setRowBGColor(46, '#E6FFE6');	
$table->setRowBGColor(47, '#E6FFE6');	
$table->setRowBGColor(48, '#E6FFE6');	
$table->setRowBGColor(49, '#E6FFE6');	
$table->setRowBGColor(50, '#E6FFE6');	
$table->setRowBGColor(51, '#E6FFE6');	
$table->setRowBGColor(52, '#E6FFE6');	
$table->setRowBGColor(53, '#E6FFE6');	
$table->setRowBGColor(54, '#E6FFE6');	
$table->setRowBGColor(55, '#E6FFE6');	
$table->setRowBGColor(56, '#E6FFE6');	
$table->setRowBGColor(57, '#E6FFE6');	
$table->setRowBGColor(58, '#E6FFE6');	
$table->setRowBGColor(59, '#E6FFE6');
$table->setRowBGColor(60, '#E6FFE6');
$table->setRowBGColor(61, '#E6FFE6');	
$table->setRowBGColor(62, '#E6FFE6');	
$table->setRowBGColor(63, '#E6FFE6');	
$table->setRowBGColor(64, '#E6FFE6');	
$table->setRowBGColor(65, '#E6FFE6');	
$table->setRowBGColor(66, '#E6FFE6');	
$table->setRowBGColor(67, '#E6FFE6');	
$table->setRowBGColor(68, '#E6FFE6');	
$table->setRowBGColor(69, '#E6FFE6');
$table->setRowBGColor(70, '#E6FFE6');
$table->setRowBGColor(71, '#E6FFE6');	
$table->setRowBGColor(72, '#E6FFE6');
$table->setRowBGColor(73, '#E6FFE6');
$table->setRowBGColor(74, '#E6FFE6');
$table->setRowBGColor(75, '#E6FFE6');
$table->setRowBGColor(76, '#87CEEB');




$table->setColAlign(2, 'right');
$table->setColAlign(3, 'right');
$table->setColAlign(4, 'right');
$table->setColAlign(5, 'right');
$table->setColAlign(6, 'right');
$table->setColAlign(7, 'right');

$table->setCellAlign(1, 1, 'center');
$table->setCellAlign(1, 2, 'center');
$table->setCellAlign(1, 3, 'center');
$table->setCellAlign(1, 5, 'center');
$table->setCellAlign(1, 6, 'center');

#$table->setCellAlign(2, 1, 'center');
#$table->setCellAlign(2, 2, 'center');
#$table->setCellAlign(2, 3, 'center');
#$table->setCellAlign(2, 4, 'center');
#$table->setCellAlign(2, 5, 'center');
#$table->setCellAlign(2, 6, 'center');
#$table->setCellAlign(2, 7, 'center');

$table->setCellAlign(2, 2, 'left');
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



$to = 'arthur.emmanuel@metroretail.com.ph,frank.gaisano@metroretail.com.ph,chit.lazaro@metroretail.com.ph, fili.mercado@metroretail.com.ph, karan.malani@metroretail.com.ph, lia.chipeco@metroretail.com.ph, marlita.portes@metroretail.com.ph, jennifer.yu@metroretail.com.ph, april.agapito@metroretail.com.ph, edna.prieto@metroretail.com.ph, tessie.baldezamo@metroretail.com.ph, chedie.lim@metroretail.com.ph, liberato.rodriguez@metroretail.com.ph,luz.bitang@metroretail.com.ph, emily.silverio@metroretail.com.ph, julie.montano@metroretail.com.ph, limuel.ulanday@metroretail.com.ph,delia.jakosalem@metroretail.com.ph,rene.babylonia@metroretail.com.ph, arthur.emmanuel@metroretail.com.ph,jayson.angeles@metroretail.com.ph,glenda.navares@metroretail.com.ph,may.sasedor@metroretail.com.ph,roy.igot@metroretail.com.ph,harvey.ong@metroretail.com.ph';


$cc = 'rex.cabanilla@metroretail.com.ph, annalyn.conde@metroretail.com.ph,roel.gevana@metroretail.com.ph,bernadette.rosell@metroretail.com.ph,fe.botero@metroretail.com.ph,jeannie.demecillo@metroretail.com.ph,mariegrace.ong@metroretail.com.ph,tessie.cabanero@metroretail.com.ph,joyce.mirabueno@metroretail.com.ph,zenda.mangabon@metroretail.com.ph,jennifer.nardo@metroretail.com.ph,liberato.rodriguez@metroretail.com.ph,eric.molina@metroretail.com.ph,rashel.legaspi@metroretail.com.ph,lanie.danong@metroretail.com.ph';



$bcc = 'lea.gonzaga@metroretail.com.ph, philip.coronado@metroretail.com.ph,dax.granados@metroretail.com.ph';


#$bcc = 'lea.gonzaga@metroretail.com.ph,eric.molina@metroretail.com.ph';

print "WHAT'S UP DOC?\n";

$from = 'Report Mailer<report.mailer@metroretail.com.ph>';

$subject = 'Sidewalk Sale:(By Merch) Hourly Sales Performance as of ' . $update_time;

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
<b><font size="4">Sidewalk Sale:Hourly Sales Report(By Merchandise) </font></b>
<br>As of &nbsp;$update_time<br><br>

$table1
$table

<p><i><font size="2">in 000s</font></i>.</p><br>

Regards, <br>
<a href= "mailto:arcbi.support@metroretail.com.ph">ARC BI Support</a>
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

