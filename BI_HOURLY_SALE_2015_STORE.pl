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
SELECT SUBSTR(STORE,2,5) AS STORE, STORE_NAME, ACTUAL, BUDGET, ACH FROM
(select 'SSTORE' STORE, 'STORE NAME' STORE_NAME, 'ACTUAL' ACTUAL, 'LAST YEAR' BUDGET, 'ACH' ACH FROM DUAL
UNION ALL
select to_char(store) store, store_name, to_char(actual, '9G999G999G999') actual, 
to_char(budget, '9G999G999G999') budget, to_char(case when actual <> 0 and budget <> 0 then round((actual / budget) * 100, 1)  else 0 end, '9G999G999G999D9') || '%' ach
from
(select a.id_str_rt store, e.store_name, 
round(sum(case when a.mo_sls_tot is not null then a.mo_sls_tot / 1000 else 0.00 end), 2) actual,
(case when b.budget is not null then round(b.budget / 1000, 2) else 0.00 end) budget
from mg_hourly_sales a, (select * from MG_BUDGET_STORE_BI where b_date = TO_CHAR(SYSDATE, 'DD-MON-YY')) b, deps c, groups d, 
(select store, store_name from store where store_name not like '%Dummy%') e
where b.store(+) = to_number(a.id_str_rt)
and to_number(a.id_str_rt) = e.store
and a.id_dpt_pos = c.dept
and c.group_no = d.group_no
and a.DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD'))
--AND a.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
--AND a.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
and d.division NOT IN ('7500', '4000', '8000', '9000', '8500')
group by id_str_rt, e.store_name, (case when b.budget is not null then round(b.budget / 1000, 2) else 0.00 end)
order by 1))
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

$table->setRowBGColor(1, '#C0C0C0');

for($i = 1; $i <= $rowcount; $i = $i + 1) {
	$table->setRowBGColor($i, '#C0C0C0');
}




my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject ) = @ARGV;




$to = 'arthur.emmanuel@metroretail.com.ph,frank.gaisano@metroretail.com.ph,chit.lazaro@metrogaisano.com, fili.mercado@metrogaisano.com, karan.malani@metrogaisano.com, luz.bitang@metroretail.com.ph,emily.silverio@metroretail.com.ph,julie.montano@metroretail.com.ph,glenda.navares@metroretail.com.ph,may.sasedor@metroretail.com.ph,roy.igot@metroretail.com.ph,manuel.degamo@metroretail.com.ph,cristy.sy@metroretail.com.ph,limuel.ulanday@metroretail.com.ph,harvey.ong@metroretail.com.ph';

$bcc = ' rex.cabanilla@metrogaisano.com, lea.gonzaga@metrogaisano.com dax.granados@metroretail.com.ph ,eric.molina@metroretail.com.ph, annalyn.conde@metroretail.com.ph,philip.coronado@metroretail.com.ph';

#$cc = 'rex.cabanilla@metrogaisano.com, annalyn.conde@metrogaisano.com,roel.gevana@metrogaisano.com,wendell.manumbaga@metrogaisano.com,bernadette.rosell@metroretail.com.ph,fe.botero@metroretail.com.ph,jeannie.demecillo@metroretail.com.ph,mariegrace.ong@metroretail.com.ph,tessie.cabanero@metroretail.com.ph,joyce.mirabueno@metroretail.com.ph,zenda.mangabon@metroretail.com.ph,jennifer.nardo@metroretail.com.ph,liberato.rodriguez@metroretail.com.ph';



#$bcc = 'lea.gonzaga@metrogaisano.com,cham.burgos@metrogaisano.com, philip.coronado@metrogaisano.com';

#$bcc = ' lea.gonzaga@metrogaisano.com,cham.burgos@metrogaisano.com';
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
<b><font size="4">Hourly Sales Report (By Store) </font></b>
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

print "rowcount: $rowcount";
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
SELECT SUBSTR(STORE,2,5) AS STORE, STORE_NAME, ACTUAL, BUDGET, ACH FROM
(select 'SSTORE' STORE, 'STORE NAME' STORE_NAME, 'ACTUAL' ACTUAL, 'LAST YEAR' BUDGET, 'ACH' ACH FROM DUAL
UNION ALL
select to_char(store) store, store_name, to_char(actual, '9G999G999G999') actual, 
to_char(budget, '9G999G999G999') budget, to_char(case when actual <> 0 and budget <> 0 then round((actual / budget) * 100, 1)  else 0 end, '9G999G999G999D9') || '%' ach
from
(select a.id_str_rt store, e.store_name, 
round(sum(case when a.mo_sls_tot is not null then a.mo_sls_tot / 1000 else 0.00 end), 2) actual,
(case when b.budget is not null then round(b.budget / 1000, 2) else 0.00 end) budget
from mg_hourly_sales a, (select * from MG_BUDGET_STORE_BI where b_date = TO_CHAR(SYSDATE, 'DD-MON-YY')) b, deps c, groups d, 
(select store, store_name from store where store_name not like '%Dummy%') e
where b.store(+) = to_number(a.id_str_rt)
and to_number(a.id_str_rt) = e.store
and a.id_dpt_pos = c.dept
and c.group_no = d.group_no
and a.DC_DY_BSN = (TO_CHAR(SYSDATE, 'YYYY-MM-DD'))
--AND a.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
--AND a.TS_RTN_HR <= (SELECT TO_CHAR(SYSDATE, 'HH24') NEW_TIME FROM DUAL)
and d.division NOT IN ('7500', '4000', '8000', '9000', '8500')
group by id_str_rt, e.store_name, (case when b.budget is not null then round(b.budget / 1000, 2) else 0.00 end)
order by 1))
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

$table->setRowBGColor(1, '#C0C0C0');

for($i = 1; $i <= $rowcount; $i = $i + 1) {
	$table->setRowBGColor($i, '#C0C0C0');
}




my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject ) = @ARGV;




#$to = ' chit.lazaro@metrogaisano.com, fili.mercado@metrogaisano.com, karan.malani@metrogaisano.com, luz.bitang@metroretail.com.ph,emily.silverio@metroretail.com.ph,julie.montano@metroretail.com.ph,glenda.navares@metroretail.com.ph,may.sasedor@metroretail.com.ph,roy.igot@metroretail.com.ph,manuel.degamo@metroretail.com.ph,cristy.sy@metroretail.com.ph,limuel.ulanday@metroretail.com.ph';

$to = 'artemm12@aol.com,frankgaisano@gmail.com ';

$bcc = 'lgnzg87@gmail.com';



#$bcc = 'lea.gonzaga@metrogaisano.com,cham.burgos@metrogaisano.com, philip.coronado@metrogaisano.com';

#$bcc = ' lea.gonzaga@metrogaisano.com,cham.burgos@metrogaisano.com';
#$bcc = ' lea.gonzaga@metrogaisano.com';


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
<b><font size="4">Hourly Sales Report (By Store) </font></b>
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

print "rowcount: $rowcount";
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

