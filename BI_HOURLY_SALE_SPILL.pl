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
use HTML::Entities;
use HTML::Table::FromDatabase;
use CGI;
use HTML::Template;

&generate_data;
&mailer;


sub generate_data{

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

$test_query = qq{ SELECT CASE WHEN NEW_TIME > 12 THEN  TO_CHAR(NEW_TIME -12) || ':14 PM' ELSE TO_CHAR(NEW_TIME) || ':14 AM' END AS UPDATE_TIME FROM (
					SELECT MAX(TO_NUMBER(TS_RTN_HR)) NEW_TIME
					FROM MG_HOURLY_SALES WHERE DC_DY_BSN = (TO_CHAR(SYSDATE-1, 'YYYY-MM-DD'))) };

$tst_query = $dbh->prepare($test_query);
$tst_query->execute();

while ( my $x =  $tst_query->fetchrow_hashref()){
	$update_time = $x->{UPDATE_TIME};
} 

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
open my $fh, ">", "hourly_sale.csv" or die "hourly_sale.csv: $!";

$test = qq{ 
SELECT 
DECODE(GROUPING(MERCH_GROUP_DESC)
             , 0, MERCH_GROUP_DESC
             , 1, 'TOTAL' 
             ) MERCH_GROUP_DESC
, DECODE(GROUPING(DIV_NAME) 
        , 0,  DIV_NAME
        , 1, 'TOTAL' || ' ' || MERCH_GROUP_DESC
        ) DIVISION
, TO_CHAR(ROUND((SUM(SALE_AMOUNT_COMP)/1000),0),'9G999G999G999') ACTUAL_COMP, TO_CHAR(SUM(BUDGET_COMP),'9G999G999G999') BUDGET_COMP, TO_CHAR(ROUND(((ROUND((SUM(SALE_AMOUNT_COMP)/1000),2)*100)/(SUM(BUDGET_COMP))),1),'9G999G999G999D9') || '%' ACH_COMP
, TO_CHAR(ROUND((SUM(SALE_AMOUNT_ALL)/1000),0),'9G999G999G999') ACTUAL_ALL, TO_CHAR(SUM(BUDGET_ALL),'9G999G999G999') BUDGET_ALL, TO_CHAR(ROUND(((ROUND((SUM(SALE_AMOUNT_ALL)/1000),2)*100)/(SUM(BUDGET_ALL))),1),'9G999G999G999D9') || '%' ACH_ALL
FROM
	(SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, SUM(H.MO_SLS_TOT) SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, 0 AS BUDGET_ALL 
	FROM MG_HOURLY_SALES H 
	  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
	  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
	  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
	WHERE DC_DY_BSN = (TO_CHAR(SYSDATE-1, 'YYYY-MM-DD')) AND D.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500') AND TO_NUMBER(H.ID_STR_RT) IN ('2009','2012','7176','7003','2006','7004','2007','6008','7005','2010','7300','7009','5006','5005','5004','5003','5002','5001','7173','3004','3003','3001','4003','7008','7007','7006','3006','3007','2003','7000','4002','3005','3002','2002','2001','2011','2008','7001','2004','7002','2005', '3012')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, SUM(H.MO_SLS_TOT) SALE_AMOUNT_ALL, 0 AS BUDGET_ALL 
	FROM MG_HOURLY_SALES H 
	  LEFT JOIN DEPS ON H.ID_DPT_POS = DEPS.DEPT
	  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
	  LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = D.DIVISION
	WHERE DC_DY_BSN = (TO_CHAR(SYSDATE-1, 'YYYY-MM-DD')) AND D.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, SUM(Q.BUDGET) AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, 0 AS BUDGET_ALL
	FROM MG_Q4_BUDGET_BI Q 
    LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
	WHERE TO_DATE(Q.B_DATE,'MM/DD/YY') = TO_CHAR(SYSDATE-1, 'DD-MON-YY') AND Q.TYPE = 'C' AND Q.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME
UNION ALL
	SELECT BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME, 0 AS SALE_AMOUNT_COMP, 0 AS BUDGET_COMP, 0 AS SALE_AMOUNT_ALL, SUM(Q.BUDGET) AS BUDGET_ALL
	FROM MG_Q4_BUDGET_BI Q 
    LEFT JOIN DIVISION D ON Q.DIVISION = D.DIVISION 
	  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = Q.DIVISION
	WHERE TO_DATE(Q.B_DATE,'MM/DD/YY') = TO_CHAR(SYSDATE-1, 'DD-MON-YY') AND Q.TYPE = 'A' AND Q.DIVISION NOT IN ('7500', '4000', '8000', '9000', '8500')
	GROUP BY BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, D.DIVISION, D.DIV_NAME)
GROUP BY ROLLUP(MERCH_GROUP_DESC, DIV_NAME) 
ORDER BY MERCH_GROUP_DESC, DIV_NAME
};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_uc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "hourly_sale.csv: $!";

$dbh->disconnect; 

}

sub mailerxx {

my $table = 'hourly_sale.csv';

my $dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
	or die $DBI::errstr;
	
my $CGI  = CGI->new();
my @COLS = (qw(DIVISION ACTUAL_COMP BUDGET_COMP ACH_COMP ACTUAL_ALL BUDGET_ALL ACH_ALL));

my $data = $dbh_csv->selectall_arrayref("
   SELECT @{[join(',', @COLS)]}
   FROM $table
", undef);

my $headers = [ 
   map {{
      URL  => $CGI->script_name, 
      LINK => ucfirst($_), 
   }} @COLS 
];

my $i;
my $rows = [
   map {
      my $row = $_;
      (++$i % 2)
         ? { ODD  => [ map { {VALUE => $_} } @{$row} ] }
         : { EVEN => [ map { {VALUE => $_} } @{$row} ] }
   } @{$data}
];

my $html = do { local $/; <DATA> };

my $template = HTML::Template->new(
   scalarref         => \$html,
   loop_context_vars => 1,
);

$template->param(
   HEADERS => $headers,
   ROWS    => $rows,
);

# print $CGI->header();
# print $template->output();

$dbh_csv->disconnect();

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject ) = @ARGV;


$to = ' kent.mamalias@metrogaisano.com';

$subject = 'Daily Sales Performance as of ' . $as_of;

my %mail = (
    To   => $to,
    Subject => $subject,
	'content-type' => "multipart/alternative; boundary=\"$boundary\""
);

$mail{Cc} = $cc if $cc;
$mail{Bcc} = $bcc if $bcc;

my $boundary = "====" . time . "====";

$mail{'content-type'} = qq(multipart/mixed; boundary="$boundary");

$boundary = '--'.$boundary;

$mail{body} = 
<<END_OF_BODY;
$boundary
Content-Type: text/html; charset="iso-8859-1"
Content-Transfer-Encoding: quoted-printable

print $CGI->header()
print $template->output()

$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

$sth->finish();
$dbh_csv->disconnect;

}

sub mailer {

my $table = 'hourly_sale.csv';

my $dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
	or die $DBI::errstr;
 
my $sth = $dbh_csv->prepare(qq{SELECT DIVISION, ACTUAL_COMP AS ACTUAL_C, BUDGET_COMP AS BUDGET_C, ACH_COMP AS ACH_C, ACTUAL_ALL AS ACTUAL_A, BUDGET_ALL AS BUDGET_A, ACH_ALL AS ACH_A
								FROM $table
							});

$sth->execute() or die "Failed to execute query - " . $dbh_csv->errstr;


my $table1 = new HTML::Table( -num_cols=>1, 
							-num_rows=>3, 
							-border=>0, );

my $table = HTML::Table::FromDatabase->new( -sth => $sth, 
                            -border=>0,
                            -width=>'0%',
                            -spacing=>10,
                            -padding=>2,
							);

$table1->setAlign('center');
$table1->setCell(1, 1, '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;');		
# $table1->setSectionCellWidth('thead', 0, 1, 4);
# $table1->setSectionCellWidth('thead', 0, 2, 2);
# $table1->setCellFormat(2, '<b>', '</b>');
$table1->setCell(1, 2, '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;COMP  STORES&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;');	
#$table1->setSectionCellWidth('thead', 0, 3, 2);	
# $table1->setCellFormat( 3, '<b>', '</b>');					
$table1->setCell(1, 3, '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ALL  STORES&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;');	

$table->setColAlign(2, right);
$table->setColAlign(3, right);
$table->setColAlign(4, center);
$table->setColAlign(5, right);
$table->setColAlign(6, right);
$table->setColAlign(7, center);
$table->setRowBGColor(9, '#BDBDBD');							
$table->setRowBGColor(15, '#BDBDBD');
$table->setRowBGColor(16, '#BDBDBD');

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject ) = @ARGV;

$to = ' arthur.emmanuel@metrogaisano.com, frank.gaisano@metrogaisano.com, lia.chipeco@metrogaisano.com, karan.malani@metrogaisano.com ';
$cc = ' eric.redona@metrogaisano.com, rex.cabanilla@metrogaisano.com, annalyn.conde@metrogaisano.com ';
$bcc = 'artemm12@aol.com, frankgaisano@gmail.com, kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com';

# $to = ' kent.mamalias@metrogaisano.com, fnaquines@metro.com.ph ';
#$to = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com ';

$subject = 'Hourly Sales Performance as of ' . $update_time;

my %mail = (
    To   => $to,
    Subject => $subject,
	'content-type' => "multipart/alternative; boundary=\"$boundary\""
);

$mail{Cc} = $cc if $cc;
$mail{Bcc} = $bcc if $bcc;

my $boundary = "====" . time . "====";

$mail{'content-type'} = qq(multipart/mixed; boundary="$boundary");

$boundary = '--'.$boundary;

$mail{body} = 
<<END_OF_BODY;
$boundary
Content-Type: text/html; charset="iso-8859-1"
Content-Transfer-Encoding: quoted-printable

<html>
Hourly Sales Report<br>
4th Quarter 2014<br>
As of $update_time<br><br>

$table1
$table
<br><p>

in 000s<br><p><br>

Regards, <br>
ARC BI Support <br>
arcbi.support&#64;metrogaisano.com<p>
</html>

$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

$sth->finish();
$dbh_csv->disconnect;

}

sub read_file {

my( $filename, $binmode ) = @_;
my $fh = new IO::File;
$fh->open("<".$filename) or die "Error opening $filename for reading - $!\n";
$fh->binmode if $binmode;
local $/;
<$fh>
	
}


__DATA__
<html>
<head>
<title>Test</title>
</head>

<body>
<h1>Test</h1>
<table>
<tr>
<TMPL_LOOP NAME=HEADERS>
   <th><a href="<TMPL_VAR NAME=URL>"><TMPL_VAR NAME=LINK></a></th>
</TMPL_LOOP>
</tr>
<TMPL_LOOP NAME=ROWS>
   <tr>
   <TMPL_UNLESS NAME="__ODD__">
      <TMPL_LOOP NAME=EVEN>
         <td style="background: #B3B3B3"><TMPL_VAR NAME=VALUE></td>
      </TMPL_LOOP>
   <TMPL_ELSE>
      <TMPL_LOOP NAME=ODD>
         <td style="background: #CCCCCC"><TMPL_VAR NAME=VALUE></td>
      </TMPL_LOOP>
   </TMPL_UNLESS>
   </tr>
</TMPL_LOOP>
</table>

</body>
</html> 