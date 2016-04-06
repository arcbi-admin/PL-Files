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
 
	$workbook = Excel::Writer::XLSX->new("Sales by Register (S6 as of $as_of).xlsx");
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
	
	foreach my $i ( "STORE", "REGISTER", "DEPT", "DESCRIPTION", "ITEM", "DESCRIPTION", "VENDOR_NAME", "COST", "RETAIL_PRICE", "QTY", "TOTAL_RETAIL_WITH_VAT", "CUR_INVEN" ) {
		$worksheet->write( 1, $col++, $i, $desc2 );
	}
	
	my $query = qq {
	Select A.Store  As Store,A.Till_Code As Register,A.GROUP_NO, A.GROUP_NAME, A.Item As Item,A.Item_Desc As Description,replace(A.Sup_Name,',',' ') As Vendor_Name,A.Cost As Cost,
A.Retail_Price As Retail_Price,Sum(A.Qty)As Qty,  Sum(A.Total_Retail_With_Vat) As Total_Retail_With_Vat,A.Cur_Inven As Cur_Inven  
	From 
		(
		Select Sh.Store,  Case When Length (Sh.Tran_No) >= 10 Then Substr (Sh.Tran_No, 1, 3)Else Substr (Lpad (Sh.Tran_No, 10, '0'), 1, 3) 
	End As Till_Code, GROUPS.GROUP_NO, GROUPS.GROUP_NAME, Si.Item As Item,I.Item_Desc,Su.Sup_Name,Si.Unit_Retail/Si.Qty As Retail_Price,Sum(Si.Qty) As Qty  ,Sum(Si.Unit_Retail) 
	As Total_Retail_With_Vat,  Soh.Stock_On_Hand As Cur_Inven,Soh.Av_Cost As Cost  
		From Sa_Tran_Head Sh 
			Join Sa_Tran_Item Si On Sh.Tran_Seq_No=Si.Tran_Seq_No 
			Join Item_Master I On Si.Item=I.Item   
			Join Item_Loc_Soh Soh On Si.Item=Soh.Item And Si.Store=Soh.Loc  
			Join Item_Supplier Its On Si.Item=Its.Item  
			JOIN SUPS SU ON ITS.SUPPLIER=SU.SUPPLIER
			JOIN DEPS ON I.DEPT = DEPS.DEPT 
			JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
		where trunc(SH.tran_datetime) = (SELECT TO_CHAR(SYSDATE-1,'DD MON YYYY') FROM DUAL) 
			And SH.sub_tran_type in ('SALE','RETURN','LAYCOMP') And Sh.Store In (2006) And Sh.Status='P' And Its.Primary_Supp_Ind='Y'  
		Group By Sh.Store,Sh.Tran_No , GROUPS.GROUP_NO, GROUPS.GROUP_NAME,
	Si.Item,I.Item_Desc,Su.Sup_Name, Si.Unit_Retail/Si.Qty,Soh.Stock_On_Hand,Soh.Av_Cost
		) A  
	Where Till_Code In ('069','073','074','076') 
	group by A.STORE,A.TILL_CODE,A.GROUP_NO, A.GROUP_NAME,A.ITEM,A.ITEM_DESC,A.SUP_NAME,A.COST,A.RETAIL_PRICE,A.CUR_INVEN
	};
	my $query_handle = $dbh->prepare($query);
	$query_handle->execute();
	
	while (my $y = $query_handle->fetchrow_hashref()) {
		$worksheet->write($a,0, $y->{STORE},$desc);
		$worksheet->write($a,1, $y->{REGISTER},$desc);
		$worksheet->write($a,2, $y->{GROUP_NO},$desc);
		$worksheet->write($a,3, $y->{GROUP_NAME},$desc);
		$worksheet->write($a,4, $y->{ITEM},$desc);
		$worksheet->write($a,5, $y->{DESCRIPTION},$desc);
		$worksheet->write($a,6, $y->{VENDOR_NAME},$desc);
		$worksheet->write($a,7, $y->{COST},$border1);
		$worksheet->write($a,8, $y->{RETAIL_PRICE},$border1);
		$worksheet->write($a,9, $y->{QTY},$desc);
		$worksheet->write($a,10, $y->{TOTAL_RETAIL_WITH_VAT},$border1);
		$worksheet->write($a,11, $y->{CUR_INVEN},$desc);
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

#$to = 'april.agapito@metrogaisano.com, rosemarie.saravia@metrogaisano.com, rhoda.camporedondo@metrogaisano.com, jay.aguilar@metrogaisano.com, seth.jaminal@metrogaisano.com, miscellaneous.mercha@metrogaisano.com, teena.velasco@metrogaisano.com, richard.ordillo@metrogaisano.com,ronald.dizon@metroretail.com.ph,chloy.lamasan@metroretail.com.ph';
$to = 'april.agapito@metrogaisano.com,ronnie.conde@metroretail.com.ph, melanie.aquino@metroretail.com.ph, rosemarie.saravia@metroretail.com.ph, ester.mendoza@metroretail.com.ph, nemesio.panugan@metroretail.com.ph, christine.lanohan@metroretail.com.ph, jay.aguilar@metroretail.com.ph, rhoda.camporedondo@metroretail.com.ph';
#$to = 'lia.chipeco@metroretail.com.ph ,edna.prieto@metroretail.com.ph,augosto.daria@metroretail.com.ph,richard.ordillo@metroretail.com.ph,toni.cuerquis@metroretail.com.ph';
#$cc = 'melinda.uy@metroretail.com.ph,jecil.cumayas@metroretail.com.ph,maruela.repuela@metroretail.com.ph,dulce.labus@metroretail.com.ph,consorcia.mullon@metroretail.com.ph,mecelle.quimbo@metroretail.com.ph,vilma.paner@metroretail.com.ph';

$bcc = 'lea.gonzaga@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'Sales by Register (S6 as of ' .$as_of. ') ';

$msgbody_file = 'message.txt';

$attachment_file = "Sales by Register (S6 as of $as_of).xlsx";

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





