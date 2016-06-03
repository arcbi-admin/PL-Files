START:

use DBI;
use DBD::Oracle qw(:ora_types);
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use Text::CSV_XS;
#use DateKey_ARC;
#use DBConnector;
use Win32::Job;
use Getopt::Long;
use IO::File;
use MIME::QuotedPrint;
use MIME::Base64;
use Mail::Sendmail;
use Date::Calc qw( Today Add_Delta_Days Month_to_Text);

($year,$month,$day) = Today();
$month_to_text = Month_to_Text($month);

##on apr 7, jerson provided a new list of items

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
	
	$workbook = Excel::Writer::XLSX->new("INSTOCK_REPORT_INTERNATIONAL_v2.xlsx");
	$bold = $workbook->add_format( bold => 1, size => 14 );
	$bold1 = $workbook->add_format( bold => 1, size => 16 );
	$script = $workbook->add_format( size => 8, italic => 1 );
	$bold2 = $workbook->add_format( size => 11 );
	$border1 = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3 );
	$border2 = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', rotation => 90, text_wrap =>1, size => 10, shrink => 1 );
	$code = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10 );
	$desc = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10 );
	$ponkan = $workbook->set_custom_color( 53, 254, 238, 230);
	$abo = $workbook->set_custom_color( 16, 220, 218, 219);
	$sky = $workbook->set_custom_color( 12, 205, 225, 255);
	$pula = $workbook->set_custom_color( 10, 255, 189, 189);
	$lumot = $workbook->set_custom_color( 17, 196, 189, 151);
	$comp = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10,  bg_color => $lumot, bold => 1 );
	$all = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10,  bg_color => $abo, bold => 1 );
	$headN = $workbook->add_format( border => 1, align => 'center', valign => 'center', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
	$headD = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 9, bg_color => $abo, bold => 1 );
	$headPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
	$headNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3, bg_color => $abo, bold => 1 );
	$headNumber = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 3, bg_color => $abo, bold => 1 );
	$head = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
	$subhead = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, bg_color => $ponkan, bold => 1 );
	$bodyN = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
	$bodyPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
	$bodyNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 3,  bold => 1);
	$body = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
	$subt = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 9);
	$down = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => 9, bg_color => $pula );

	# printf "Test ETL Status = ". $test ." \nArc BI Sales Performance Part 1\nGenerating Data from Source \n";
	
	# $tst_query->finish();
	
	&generate_csv;
	
	&new_sheet($sheet = "Store");
	
	&call_summary_division($area_flg1 = 3, $area_flg2 = 3);
	
	&new_sheet($sheet = "Warehouse");
	
	&call_summary_division($area_flg1 = 1, $area_flg2 = 2);
	
	$workbook->close();
	$dbh_csv->disconnect;
	
	&mail1;
	&mail2;
	&mail3;
	
	exit;
	
}

else{
	print "Test Status = ". $test ." \nETL still running\n";
	
	$tst_query->finish();
	
	sleep(600);
	
	goto START;
}
 
 
sub call_summary_division {

##=============China
$a = 0; $vis = 7; $col = 7; $test = 1; $stopper = 0; 

&heading;

$country = 'China_HK_Taiwan', $region = 'China/HK/Taiwan', $a += 6, $counter = 0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $stopper = 1, $col += 3;

##==========INDIA

$country = 'India', $region = 'India', $a += 6, $counter  =0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

##==========INDONESIA

$country = 'Indonesia', $region = 'Indonesia', $a += 6, $counter  =0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

##==========JAPAN

$country = 'Japan', $region = 'Japan', $a += 6, $counter  =0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;


##==========KOREA

$country = 'Korea', $region = 'Korea', $a += 6, $counter  =0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

##==========MALAYSIA

$country = 'Malaysia', $region = 'Malaysia', $a += 6, $counter  =0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

##==========SINGAPORE

$country = 'Singapore', $region = 'Singapore', $a += 6, $counter  =0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

##==========THAILAND

$country = 'Thailand', $region = 'Thailand', $a += 6, $counter  =0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

##==========VIETNAM

$country = 'Vietnam', $region = 'Vietnam', $a += 6, $counter  =0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

$worksheet->set_column( 7, 8, undef, undef, 1 );
$worksheet->set_column( 10, 11, undef, undef, 1 );
$worksheet->set_column( 13, 14, undef, undef, 1 );
$worksheet->set_column( 16, 17, undef, undef, 1 );
$worksheet->set_column( 19, 20, undef, undef, 1 );
$worksheet->set_column( 22, 23, undef, undef, 1 );
$worksheet->set_column( 25, 26, undef, undef, 1 );
$worksheet->set_column( 28, 29, undef, undef, 1 );
$worksheet->set_column( 31, 32, undef, undef, 1 );

}



sub new_sheet{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(85);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
$worksheet->set_print_scale( 100 );
$worksheet->set_margins( 0.05 );
$worksheet->conditional_formatting( 'F9:V2000',  { type => 'cell', criteria => '<', value => 0, format => $down } );
$worksheet->set_column( 1, 3, undef, undef, 1 );
# $worksheet->set_column( 1, 2, 3 );
# $worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 21 );

}

sub heading {

$worksheet->write(0, 4, "Replenishment In-Stock " . $comment , $bold);
$worksheet->write(1, 4, $day . '-' . $month_to_text . '-' .$year, $bold2);

$worksheet->merge_range( 4, 2, 5, 2, 'Type', $subhead );
$worksheet->merge_range( 4, 3, 5, 3, 'Type', $subhead );
# $worksheet->merge_range( 4, 4, 5, 4, 'Type', $subhead );
# $worksheet->merge_range( 4, 5, 5, 5, 'Code', $subhead );
$worksheet->merge_range( 4, 4, 5, 6, 'Location', $subhead );

}
 

sub query_summary_division {

$table = 'instock_asian.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls1 = $dbh_csv->prepare (qq{SELECT country, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
									 FROM $table 
									 WHERE country = '$country' and (area <> '$area_flg1' and area <> '$area_flg2')
								GROUP BY country
							 }); $sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$country = $s->{country};
	
	$sls2 = $dbh_csv->prepare (qq{SELECT area, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
									 FROM $table 
									 WHERE country = '$country' and (area <> '$area_flg1' and area <> '$area_flg2')
									 GROUP BY area
									 ORDER BY area
								}); $sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$area = $s->{area};
				
		$sls3 = $dbh_csv->prepare (qq{SELECT store, store_name, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
									 FROM $table 
									 WHERE country = '$country' AND area = '$area'
									 GROUP BY store, store_name
									 ORDER BY store
									}); $sls3->execute();
									
		while(my $s = $sls3->fetchrow_hashref()){
		
			if($stopper eq 0){ $worksheet->merge_range( $a, 4, $a, 6, $s->{store}.'-'.$s->{store_name}, $desc );}
			
			$worksheet->write($a,$col, $s->{tot_items},$border1); 
			$worksheet->write($a,$col+1, $s->{rep_items},$border1);
				if ($s->{tot_items} <= 0){
					$worksheet->write($a,$col+2, "",$subt); 				}
				else{
					$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt); 				}
			
			$counter = 0; #RESET dept_counter	
			$a++; #INCREMENT VARIABLE a
		}

		$worksheet->write($a,$col, $s->{tot_items},$bodyNum); 
		$worksheet->write($a,$col+1, $s->{rep_items},$bodyNum);
		if ($s->{tot_items} <= 0){
			$worksheet->write($a,$col+2, "",$bodyPct); 		}
		else{
			$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$bodyPct); 		}
		
		if($stopper eq 0){
			$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );}
		
		$a++; #INCREMENT VARIABLE a
	}
	
	$worksheet->write($a,$col, $s->{tot_items},$headNumber); 
	$worksheet->write($a,$col+1, $s->{rep_items},$headNumber);
	if ($s->{tot_items} <= 0){
		$worksheet->write($a,$col+2, "",$headPct); 	}
	else{
		$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$headPct); 	}

}

if($stopper eq 0){		 	$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN ); }

$worksheet->merge_range( $loc_pt, $col, $loc_pt, $col+2, $region, $subhead );

$sls1->finish();
$sls2->finish();
$sls3->finish();
#$sls4->finish();

$len = $a, $counter = 0;

}



sub generate_csv {

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
# my $uid = 'kent';
# my $pw = 'amer1c8';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh, ">", "instock_asian.csv" or die "instock_asian.csv: $!";

$test = qq{ 
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'China_HK_Taiwan' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (3010,6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
			WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
		FROM(
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('805017','805033','1985535','1998153','1998160','1998450','1998535','1998542','2015859','2016276','2016528','2034379','2034409','2034447','2034485','2034522','2058832','2058894','2071831','2077420','3714188','3771808','3782675','3782682','3814468','3816592','5231157','5250516','5251544','5292912','5293810','5405107','5469208','5485345','5560868','5650866','5681877','5702169','5702305','5731602','5735525','5804399','5804504','5804665','5805006','5805488','5807499','5842841','5915231','6012069','6018511','6035396','6035419','6035457','6064310','6214050','6214395','6279073','6279080','6279110','6312442','6415419','6442606','6442804','6444655','6444686','6461751','6461768','6461775','6543075','7295614','7296482','7296574','7296604','7456138','7460845','7466359','7525209','7525261','7643187','7746970','7749285','7820014','7830440','7919183','8022066','8031389','8221254','8404824','8432575','8470973','8472793','8472816','8472823','8472830','8628244','8628336','8628787','8628800','8628817','8628848','8647399','8847652','9018150','9026080','9031411','9031412','9031413','9032002','9035417','9100037','9123909','9123911','9128701','9139545','9139558','9147310','9147314','9147321','9147324','9147334','9147336','9147654','9147659','9147662','9147663','9150556','9150559','9150565','9165883','9168091','9177495','9178372','9178384','9178629','9192578','9218312','9244712','9244725','9249726','9253621','9253658','9254263','9255919','9266594','9286293','9326420','9337470','9337518','9337524','9357773','9358814','9358845','9358862','9358864','9358870','9360401','9360402','9360412','9360414','9360415','9364502','9365934','9367158','9367162','9367528','9368842','9370048','9370055','9373176','9398075','9404468','9405497','9406493','9409395','9416853','9416854','9416856','9416857','9416858','9416865','9416868','9416874','9422784','9425482','9430028','9456779','9472548','9478331','9478336','9478342','9478348','9478372','9478374','9478379','9478381','9478384','9482286','9482287','9482288','9482289','9482290','9482291','9482398','9482406','9505460','9529773','9529774','9529775','9529777','9529780','9529783','9534202','9539624','9539808','9558071','9580895','9580956','9581234','9581256','9581257','9581484','9584200','9590139','9590155','9591253','9591254','9598624','9602313','9614840','9614843','9614844','9614847','9615239','9623025','9639722','9639728','9664669','9664670','9666691','9667237','9668220','9670804','9678355','9678521','9695858','9697906','9697964','9698047','9698054','9698182','9698247','9705188','9708884','9708894','9738576','9738586','9738591','9738597','9738604','9748197','9748385','9748388','9748389','9748852','9748854','9748857','9748864','9748865','9748867','9748873','9757236','9762054','9762061','9762064','9779126','9790966','9795326','9795327','9795330','9795332','9795333','9795334','9795989','9795991','9796029','9796046','9796090','9796104','9796108','9798375','9798393','9798470','9798475','9801001','9801004','9801014','9816565','9830920','9830955','9830970','9830976','9830979','9831040','9834299','9834318','9834333','9834342','9834347','9834353','9834427','9834483','9834504','9834510','9834522','9835201','9835214','9835980','9840920','9840921','9840922','9843918','9843923','9850263','9862689','9862690','9863370','9864003','9864005','9864007','9864008','9864010','9864302','9864305','9873134','9873272','9873304','9873450','9873461','9873760','9874116','9874148','9876754','9877844','9877853','9877870','9884149','9885266','9887713','9887715','9889408','9892424','9892426','9892435','9892437','9892439','9892445','9892446','9892469','9892474','9892476','9892490','9893797','9894805','9897618','9902430','9903859','9908555','9908705','9908710','9913286','9913296','9913314','9915894','9917977','9917979','9923533','9930844','9930853','9930887','9934299','9934305','9934321','9934324','9934328','9934334','9934346','9934349','9935890','9947332','9964569','9980628','9980640','9985497','9986726','9986728','9986765','9986766','9992327','9996626','9996632','9996636','9996657','10004494','10004496','10004498','10004499','10004500','10004501','10004502','10004503','10004504','10004505','10004507','10004508','10004509','10004510','10004511','10004512','10004513','10004514','10004516','10004521','10004528','10004535','10004541','10004554','10004560','10004564','10004566','10004568','10004570','10004571','10004572','10004596','10004597','10004598','10004600','10004608','10004610','10004612','10004620','10004622','10004625','10004632','10004633','10004635','10004636','10004637','10004674','10004677','10004680','10004683','10004685','10004688','10004692','10004693','10004694','10004696','10004697','10004698','10004699','10004700','10004701','10004702','10004704','10004705','10004707','10004708','10004710','10004711','10004715','10004716','10004718','10004720','10004723','10004725','10004733','10004734','10004735','10004739','10004741','10004744','10004752','10004757','10004839','10004899','10004915','10004930','10004936','10004938','10004942','10004953','10004959','10004964','10004967','10005000','10005001','10005002','10005007','10005008','10005009','10005011','10005012','10005013','10005016','10005018','10005020','10005021','10005057','10005061','10005082','10005125','10005130','10005138','10005142','10005145','10005160','10005164','10005165','10005342','10005346','10005366','10005369','10005371','10005373','10005375','10005381','10005383','10005385','10005390','10005392','10005486','10005493','10005504','10005574','10005579','10005605','10005608','10005613','10006367','10006387','10006392','10006535','10006544','10006548','10006549','10006583','10006588','10007260','10007847','10008309','10010602','10010604','10010962','10011154','10012692','10015242','10032102','10032103','10032104','10032105','10032106','10032107','10032108','10045108','10052750','10056485','10083415','10083419','10087459','10088256','10088312','10088914','10088919','10088922','10089245','10089266','10089273','10089274','10089275','10089276','10089278','10089281','10089588','10089594','10089597','10089599','10089603','10089607','10089623','10089626','10089632','10092678','10092679','10092681','10092683','10092684','10092686','10092687','10092689','10092690','10092691','10092692','10092693','10092695','10092697','10092698','10092700','10092701','10092703','10092704','10092705','10092707','10092709','10092710','10092715','10092717','10094258','10094271','10094826','10095676','10095941','10095944','10101779','10101934','10101935','10101936','10101937','10101938','10102330','10103674','10103675','10104477','10104505','10104513','10104515','10104516','10104630','10104633','10104636','10104638','10104640','10104642','10104643','10104644','10105959','10108123','10108133','10121005','10122074','10123587','10123588','10147540','10151471','10151486','10151488','10151490','10151492','10151709','10151717','10151721','10151722','10153626','10153684','10153688','10153694','10153696','10153697','10153699','10153701','10158327','10158685','10161491','10163146','10164685','10164692','10164698','10164701','10164704','10164707','10165728','10165730','10168588','10168590','10170209','10170214','10172399','10174582','10174606','10174923','10175210','10175219','10175223','10175230','10175234','10175235','10175236','10175237','10175238','10175239','10175263','10175272','10175287','10175290','10175291','10175295','10175298','10175301','10175304','10175305','10175307','10175309','10175406','10175418','10177176','10181499','10181504','10181507','10181773','10186071','10187105','10187113','10187168','10187173','10187174','10187186','10187191','10187197','10187205','10187206','10187207','10187222','10187223','10187321','10187325','10187801','10187806','10187808','10187814','10187816','10187818','10188446','10192020','10193552','10193557','10193561','10193575','10193576','10193578','10193580','10193581','10193586','10193588','10193596','10193601','10193602','10193606','10195047','10195055','10195057','10195060','10195061','10195063','10195068','10195070','10195073','10200853','10200858','10200862','10201817','10201826','10202036','10202046','10202048','10202117','10202121','10202123','10202124','10202133','10202136','10202139','10202344','10202345','10202356','10202372','10202583','10202595','10202598','10202884','10202891','10202898','10202901','10202928','10202930','10202936','10202941','10202944','10202947','10202955','10202973','10202980','10202989','10202994','10202998','10203001','10203003','10203005','10203007','10203009','10203011','10203013','10203016','10203017','10203329','10203347','10203351','10203358','10203364','10203368','10203373','10203378','10203380','10203392','10203394','10203395','10203399','10203402','10203403','10203404','10203405','10203407','10203471','10215366','10215373','10215381','10215396','10216709','10216715','10216725','10216726','10216727','10216735','10216746','10216755','10216763','10216770','10216780','10216804','10216805','10216806','10216808','10216810','10216813','10216814','10216815','10217463','10217470','10224521','10224998','10224999','10225032','10225033','10225093','10225098','10225100','10225104','10229130','10229428','10229732','10233376','10233400','10233402','10239970','10239978','10244993','10245016','10245025','10245321','10245345','10245366','10245392','10245429','10246715','10246721','10247890','10247893','10247894','10247898','10247899','10247900','10247903','10247904','10249383','10255754','10255773','10255782','10255818','10255847','10256233','10256258','10256301','10256441','10256569','10256578','10256605','10256612','10256618','10256624','10256628','10256632','10256641','10256648','10256654','10256721','10256729','10256741','10256745','10256758','10257307','10261352','10261369','10261374','10261384','10261388','10261391','10261448','10261449','10261450','10261452','10304890','10304921','10304948','10304973','10304999','10305024','10305039','10305047','10305054','10305055','10305056','10305057','10305058','10305059','10305060','10305061','10305062','10305063','10305064','10305065')
			UNION ALL
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('10305066','10305067','10305068','10305069','10305070','10305071','10305072','10305074','10305075','10305076','10305077','10305079','10305080','10305081','10305178','10305179','10305193','10305195','10305196','10305216','10305223','10305230','10305234','10305237','10305239','10305240','10305242','10305244','10305247','10305248','10305249','10305251','10305253','10305255','10305259','10305260','10305318','10305319','10305322','10305325','10305327','10305328','10305329','10305330','10305331','10305332','10305333','10305334','10305336','10305339','10305342','10305343','10305344','10305346','10305347','10305350','10305351','10305359','10305367','10305383','10305392','10305393','10305395','10305396','10305397','10305398','10305399','10305400','10305401','10305402','10305403','10305404','10305405','10305406','10305407','10305408','10305409','10305410','10305411','10305412','10305413','10305414','10305415','10305416','10305417','10305418','10305419','10305420','10305421','10305422','10305423','10305424','10305425','10305426','10305427','10305428','10305429','10305430','10305431','10305432','10305433','10305434','10305435','10305436','10306168','10306204','10306240','10306292','10306296','10306422','10306425','10306432','10306484','10306499','10306528','10306530','10306531','10306536','10306549','10306555','10308729','10309749','10309752','10309755','10309757','10309813','10309848','10309865','10309881','10309898','10309974','10309977','10309980','10309986','10309991','10309995','10309998','10310001','10310003','10310006','10310008','10310013','10310015','10310022','10310041','10310061','10310080','10310099','10310114','10310122','10310134','10310168','10310170','10310172','10310173','10310177','10310568','10310587','10310715','10310733','10310756','10310760','10310843','10313391','10313393','10313395','10313397','10313443','10313475','10313476','10313477','10313478','10313479','10313480','10313483','10313485','10313487','10313497','10313500','10313506','10313508','10313512','10313515','10313518','10313540','10313566','10313569','10313572','10313573','10313582','10313588','10313652','10313653','10313654','10313993','10313997','10313999','10314001','10314004','10314046','10314091','10314092','10314105','10314128','10314156','10314179','10314203','10314233','10314254','10314301','10314304','10314306','10314307','10314309','10314311','10314348','10314349','10314350','10314351','10314353','10314354','10314358','10314361','10314364','10314366','10314370','10314371','10314372','10314374','10314376','10314378','10314379','10314380','10314381','10314382','10314383','10314391','10314392','10314394','10314396','10314397','10314399','10314414','10314454','10314470','10314767','10314776','10315068','10315132','10315152','10316177','10316215','10316235','10316266','10316282','10317882','10317900','10317908','10317912','10317915','10317916','10317919','10317921','10317922','10317923','10317925','10317926','10317928','10317959','10317960','10317961','10318040','10318076','10318111','10318138','10318161','10318169','10318175','10318182','10318234','10318304','10318319','10318338','10318362','10318388','10318408','10318651','10318665','10318670','10318676','10318680','10318683','10318689','10318693','10318860','10318867','10318928','10318956','10318961','10318968','10318973','10318976','10319010','10319032','10319126','10319134','10319139','10319143','10319147','10319150','10319157','10319253','10319274','10319302','10319303','10319312','10319318','10319324','10319333','10319347','10319350','10319355','10319443','10319458','10319460','10319467','10319486','10319487','10319490','10319493','10319517','10319535','10319576','10319591','10319613','10319624','10319635','10319661','10319673','10319713','10319807','10319809','10319812','10319816','10323764','10325059','10325060','10325061','10329502','10333035','10335888','10340854','10340857','10340859','10340862','10340863','10340874','10340884','10340886','10340888','10348000','10350985','10350999','10351013','10351041','10366167','10369118','10369126','10375679','10377030','10389575','10389588','10389602','10411502','10411695','10412742','10412847','10412881','10413040','10413063','10413106','10424657','10424658','10424660','10424661','10424662','10424663','10424664','10424665','10424666','10424667','10424668','10424669','10424673','10424674','10430761','10430762','10430764','10430765','10430766','10430767','10430788','10430813','10430829','10432922','10437450','10438715','10438843','10438844','10439861','10440027','10440045','10440065','10445863','10445864','10450878','10470705','10470706','10476800','10476803','10477067','10481140','10481141','10481142','10483603','10483604','10490222','10490223','10490224','10490225','10490226','10490227','10490228','10491394','10491597','10491599','10495319','10496977','10496978','10511651','10511652','10511653','10523719','10523721','10523723','10523724','10523725','10523726','10523727','10523728','10523729','10523731','10523732','10523740','10523742','10523744','10523746','10523748','10523750','10523752','10523754','10523756','10523758','10523759','10523760','10523761','10523762','10523764','10523765','10523766','10523767','10523768','10523769','10523770','10523771','10523772','10523773','10523774','10523775','10523776','10523777','10523778','10523779','10523780','10523782','10523783','10523785','10523786','10523787','10523788','10523789','10523790','10523791','10523792','10523793','10523794','10523795','10523796','10523797','10523799','10523800','10523801','10523802','10523803','10523804','10523805','10523806','10523807','10523808','10523809','10524207','10524330','10531925','10531926','10531927','10531928','10531929','10533542','10533543','10533544','10533545','10533546','10533547','10533548','10533551','10535460','10535462','10535464','10552522','10552523','10552972','10552973','10552974','10552976','10552977','10555430','10557129','10557131','10557132','10569818','10569819','10569820','10569821','10569823','10569825','10569827','10569829','10569830','10569831','10569834','10569835','10569836','10569838','10569839','10569841','10590909','10590911','10590914','10590917','10599966','10604284','10604285','10604286','10604287','10604289','10604290','10604291','10604292','10604293','10604294','10604295','10604302','10604316','10604331','10608015','10608017','10608019','10608021','10608023','10608043','10608044','10608045','10608046','10608047','10608048','10608049','10608050','10608051','10608052','10608053','10608054','10608055','10608056','10608057','10608058','10608059','10608060','10608061','10608062','10608063','10608064','10608065','10608066','10608067','10608068','10608091','10608092','10608093','10608094','10608095','10608096','10608097','10608099','10608102','10608104','10608106','10608108','10608111','10608113','10608115','10613571','10613572','10613573','10613574','10613575','10613577','10613580','10613582','10613585','10613587','10621962','10621963','10621964','10621965','10621966','10624082','10624083','10627029','10627030','10659893','10659894','10659896','10661963','10661966','10661967','10661968','10661969','10666453','10666454','10666455','10666456','10679695','10679696','10679697','10688672','10688678','10688679','10688680','10688681','10688682','10688683','10688684','10688685','10688686','10689503','10689506','10689508','10689510','10690282','10690283','10694504','10694505','10694506','10694530','10694540','10698344','10698345','10698346','10698347','10705180','10709741','10709742','10709743','10709744','10709745','10709746','10709747','10709748','10709749','10709750','10709751','10709752','10709753','10709754','10709755','10709756','10709757','10709758','10709759','10738166','10738167','10738168','10738173','10738174','10738175','10746548','10756624','10756628','10763473','10783859','10783860','10783861','15001213','15001218','15001391','15001392','15002432','15002762','15002873','15003109','15003110','15003152','15003237','15003696','101443626','101506255','101506276','101506345','101506604','101506645','101506868','101506886','101506899','101506942','101506982','101507004','101507024','101507042','101507057','101507084','101507140','101507170','101507205','101507257','101507315','101507373','101507403','101507488','101507541','101507696','101507746','101513888','101541830','101541839','101541884','101541888','101541896','101548498','101579502','101579508','101579517','101579526','101579600','101581967','101581982','101582057','101582069','101582094','101582440','101582458','101582466','101582905','101582914','101582937','101582949','101583035','101583981','101583983','101585201','101585274','101591981','101591986','101591991','101620820','101632795','101632798','101632801','101632804','101646770','101646778','101646781','101646784','101646790','101646802','101646808','101646831','101646845','101646849','101647207','101647213','101647239','101647244','9798625','10313987','10558519','10657108','9817186','9874083','9874085','9892395','9892403','9892404','9892405','10099116','10099117','10313708','10318995','10319017','10319026','10319164','10329501','10348003','1991758','9023384','9397210','9720045','10022916','10063251','10160242','10160244','10237912','10251821','10251824','10271132','10271136','10271487','10271494','10601718','10768991','10768992','10768993','101702898','101702985','101703318','101704586','101704993','101705001','101730982','101731342','101741078')
			)TBL
	GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO
UNION ALL
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'India' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
		WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
		FROM(
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('10245315','10245323','10245344','10245401','10245445','10351160','101443652')
			)TBL
		GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO		
UNION ALL
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'Indonesia' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
		WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')	
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
		FROM(
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('6374266','9017380','9358796','9783398','9783399','9783401','9783405','9783452','9783457','9783647','10245373','10245404','10245409','10245413','10245434','10245440','10245446','10473232','10473234','10473236','10473238','10473240','10477062','10477064','10477065','10477066','10495318','10535465','10778308','10778310','101471545','101471577','101471584','101585533','101587133')
			)TBL
		GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO	
UNION ALL
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'Japan' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
		WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')	
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
		FROM(
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('3792056','5236466','5670932','5807383','6073268','6073596','6073770','6279820','7197659','7197666','8730572','8730688','9027215','9041999','9178408','9373129','9385712','9385778','9385830','9385916','9386314','9386347','9386359','9396348','9440856','9450442','9450447','9450450','9509410','9509446','9589104','9589221','9589225','9589235','9589238','9589242','9671415','9702129','9771459','9771461','9771535','9771553','9772586','9772608','9772677','9787816','9788813','9796924','9798301','9805315','9805470','9805494','9832345','9832362','9832864','9833050','9833080','9833088','9833194','9833202','9833238','9833265','9833271','9833287','9833292','9833295','9833305','9833310','9833440','9833452','9833464','9833473','9833503','9833505','9833507','9833669','9833671','9833689','9833694','9833719','9833725','9833730','9833734','9833738','9833745','9833754','9833758','9833760','9833763','9833764','9833767','9833771','9833784','9833827','9833833','9833843','9833850','9833861','9836074','9836178','9847453','9847454','9848247','9848300','9884700','9884702','9884705','9890761','9905016','9905019','9905020','9905062','9905082','9928884','9995016','9995253','10015509','10015521','10020913','10049847','10073450','10089881','10113346','10113347','10113348','10113349','10113350','10113352','10113355','10113358','10150262','10150266','10150269','10150271','10150273','10150670','10150673','10150675','10150683','10150686','10150689','10150692','10150694','10170014','10175016','10175026','10175028','10175030','10175046','10175052','10175056','10177145','10190116','10190128','10190132','10190142','10190154','10190157','10190185','10190194','10190195','10190197','10190199','10190206','10190217','10190220','10190222','10190225','10190237','10198907','10198908','10198952','10198960','10198977','10199002','10199005','10199006','10199007','10199009','10199011','10199014','10199017','10199018','10199022','10199023','10199027','10199030','10199031','10199036','10245381','10245406','10251862','10251882','10251898','10251907','10251914','10447981','10447982','10447985','10448035','10448085','10448121','10448150','10448175','10448424','10448448','10448479','10448501','10450914','10450950','10450983','10451004','10451037','10451058','10451077','10497163','10555411','10555414','10555416','10555418','10555419','10555421','10555423','10626845','10626847','10626850','10626851','10626852','10626853','10626854','10626855','10626856','10626857','10626858','10626859','10626860','10626861','10626862','10626863','10626864','10626865','10626866','10626867','10626868','10626869','10626870','10626871','10626872','10626873','10626874','10626875','10626877','10626878','10626880','10626881','10626882','10626883','10626884','10626885','10626886','10626887','10626888','10626889','10626891','10626893','10626896','10626898','10626901','10626903','10626906','10626908','10626911','10626912','10626913','10626914','10626915','10626916','10626917','10626918','10626919','10626920','10626921','10626922','10626923','10626924','10626925','10626926','10626927','10626928','10626929','10626931','10626932','10626933','10626934','10626935','10626936','10626937','10626938','10626939','10626940','10626941','10626942','10626943','10626944','10626945','10626946','10626947','10626948','10626949','10626950','10626951','10626952','10626953','10626954','10626955','10626956','10626957','10626958','10626959','10626960','10626961','10626962','10626963','10626964','10626965','10626966','10626967','10626968','10626969','10626970','10626971','10626972','10626973','10626974','10626975','10626976','10626977','10626978','10626979','10626980','10626981','10626982','10626983','10626984','10626985','10626986','10626987','10626988','10626989','10626990','10626991','10626992','10626993','10626994','10626995','10626996','10626997','10626998','10626999','10627000','10627001','10627002','10627003','10627004','10627005','10627006','10627007','10627008','10627009','10627010','10627011','10627012','10627013','10627016','10627018','10627020','10627023','10627025','10627026','10627027','10627028','10628102','10628103','10628104','10628105','10628106','10628107','10628108','10628109','10628111','10628113','10628114','10628115','10628116','10628117','10628118','10628120','10628121','10628122','10628123','10628124','10628125','10628126','10628127','10670295','10670296','10676297','10676298','10676299','10767153','10767154','10767155','10767156','10767157','10767158','10767159','10767160','10767161','10767162','10767163','10767164','10767165','10767166','10767167','10767168','10767169','10767170','10767171','10767172','10767173','10767174','10767175','10767176','10767177','10767178','10767179','10767180','10767181','10767182','10778313','15000300','15001600','15001604','101434165','101570345','101570535','101570571','101570617','101570645','101570656','101570658','101570660','101570663','101570667','101570673','101570825','101570879','101570995','101571030','101571378','101571420','101571457','101571508','101571600','101571673','101571813','101571879','101571931','101571955','101572008','101572033','101572080','101572103','101572628','101572672','101573945','101573987','101573998','101574039','101574061','101574118','101574137','101574161','101574177','101574363','101574388','101574397','101574411','101574420','101574428','101574570','101574581','101574609','101574814','101574816','101577037','101577043','101577057','101577061','101577067','101577085','101577094','101577104','101577115','101577124','101577158','101577161','101577230','101577239','101577281','101577285','101577307','101577326','101577355','101577368','101577376','101577386','101577396','101577403','101577413','101577421','101577434','101577460','101577464','101577481','101577500','101577509','101577516','101577519','101577522','101577524','101577527','101577532','101577536','101578601','101578603','101579386','101579388','101579390','101579405','101579410','101585806','101585820','101585824','101585827','101636445','101636471','101636477','101636483','9834732','9834736','9917034','101616489','9588817','9588917','9771509','9772336','9772346','9772351','9772353','9772358','9772362','9772574','9772576','9772595','9833230','9865479','9865482','9865498','9865502','9865504','9865548','9890792','9905023','9905029','9905037','9905043','9905093','9922816','10020939','10020942','10171609','10171615','101703206','101703391','101703523','101703527','101703545','101703578','101703657','101703680','101703701','101703646','101703715','101703733')
			)TBL
		GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO	
UNION ALL
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'Korea' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
	WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')		
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
		FROM(
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('7461019','9583723','9583727','9583732','9583735','9583743','9583747','9583749','9583753','9583765','9583768','9583770','9583771','9583783','9771344','9928870','10020904','10171586','10171596','10171601','10171606','10171626','10171633','10171641','10171645','10228726','10228730','10228736','10228743','10228748','10270305','10451134','10451199','10451291','10451340','10451419','10451458','10451493','10451534','10451571','10608024','10608026','10608027','10608028','10608029','10608030','10608031','10608032','10608033','10608034','10608035','10608036','10608037','10608038','10608039','10608040','10608041','10608042','10738169','10738170','10738171','101416082','101442519','101442620','101442639','101442689','101442720','101442825','101442854','101446876','8462862')
			)TBL
		GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO	
UNION ALL
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'Malaysia' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
		WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')	
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
		FROM(
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('6018405','9358795','9814419','9814420','9895753','9895764','9895769','9895774','9895775','9895777','9895779','9895782','9895784','9895789','9992196','9992197','10046345','10046359','10046371','10079856','10083470','10083485','10083492','10083495','10083501','10083514','10083528','10083532','10083537','10094395','10094415','10094420','10094426','10094430','10094432','10177626','10177652','10177659','10177668','10177670','10177674','10177678','10177682','10177686','10177692','10177695','10177697','10177701','10177707','10177712','10177717','10177718','10177720','10177723','10177726','10177727','10177729','10177743','10177754','10177762','10177770','10177775','10177783','10177793','10177799','10177804','10205065','10205074','10205083','10205104','10231029','10231041','10236871','10236879','10236884','10236888','10236898','10358108','10358121','10480264','10480265','10480266','10481144','10556641','10556642','10556644','10556647','10556649','10556651','10556654','10556965','10556966','10556967','10556968','10674085','10674086','10674087','10679102','10679103','10679104','10679105','10679106','10679107','10679108','10679109','10679110','10679111','10679112','10679113','10679114','10679115','10679116','10679117','10679118','10679119','10679120','10679121','10679122','10750271','10754937','10754940','10768235','10768236','10768238','10768239','10768240','10768242','101443875','101443893','101443953','101443967','101443990','101443998','101444070','101444119','101444131','101444142','101484822','101643165','101643203','101643207','101643448','101644626','101644648','101644666','101644683','101644692','101649642','101675619','101433958','101434020','101434038','101434056','101434096','101434137','101599149','101599151','101599158','101599164','101704839','101704847','101704857')
			)TBL
		GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO
UNION ALL
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'Singapore' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
		WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')	
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
		FROM(
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('5231133','8404909','10107301','10107303','10107314','10107316','10107317','10107329','10107345','10107350','10107363','10107366','10107369','10107373','10107382','10113292','10145380','10146100','10146101','10146164','10146186','10146195','10146210','10146220','10146223','10146244','10146254','10146256','10146257','10146260','10146264','10146271','10146286','10146294','10146297','10146299','10146311','10146335','10146337','10146338','10146339','10146343','10146346','10146370','10146392','10146446','10146460','10146486','10146490','10153498','10153508','10172488','10172491','10172492','10172495','10172505','10174697','10174699','10205691','10205827','10205829','10205836','10245339','10245418','10245423','10245439','10256522','10351089','10663604','10663605','10663607','10663609','10663612','10663614','10663616','10663618','10663620','10663622','10663624','10663648','10663649','10663650','15002435','101423440','101591713','101591845','101591856','101591865','101703373','101703381')
			)TBL
		GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO	
UNION ALL
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'Thailand' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
		WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')	
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
		FROM(
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('10510576','10510577','805072','8312501','8432551','8846389','8846419','9017283','9152370','9175521','9191874','9191875','9191879','9191883','9228737','9228738','9228739','9228742','9228743','9271184','9271199','9271200','9271209','9285459','9285465','9285468','9285471','9334966','9334970','9334973','9334975','9334988','9334992','9363025','9363055','9363073','9363075','9363105','9363147','9363153','9363154','9363157','9447967','9466230','9534497','9566974','9566975','9566982','9566985','9566987','9575545','9575561','9575565','9583118','9583119','9589587','9589600','9589610','9589612','9589615','9589620','9589624','9589630','9589631','9589681','9589696','9604844','9639388','9639392','9639430','9639672','9639676','9639678','9639680','9667460','9671873','9702131','9702132','9783506','9783508','9783509','9783510','9783511','9783528','9783550','9783558','9783562','9783566','9783567','9783581','9783586','9783587','9783588','9798806','9798809','9798842','9805863','9809965','9809968','9809972','9811382','9812238','9856189','9883272','9883279','9883286','9883291','9884322','9884330','9949635','9949644','10029621','10062312','10062328','10062429','10062435','10062442','10084085','10084094','10114459','10114460','10124897','10124899','10124902','10124905','10124906','10124907','10124911','10124914','10124921','10124925','10124930','10124935','10124938','10124946','10124950','10168758','10168760','10168761','10168762','10168764','10168765','10186117','10186118','10186119','10186120','10186121','10186122','10197510','10200431','10200439','10200444','10200451','10204839','10205660','10205666','10205671','10205708','10205711','10205825','10205831','10245331','10245355','10245433','10245437','10245438','10245443','10245447','10251795','10251796','10251797','10251798','10251799','10251802','10251803','10251804','10270334','10270454','10270525','10270540','10270578','10270613','10350542','10350547','10350553','10350562','10350569','10350573','10350574','10350578','10350605','10350637','10350662','10350686','10350930','10364389','10364392','10364394','10364480','10364502','10364503','10364531','10364532','10364771','10364784','10442678','10443197','10443267','10443268','10443269','10443270','10443271','10443272','10450117','10450127','10450172','10450195','10450207','10450807','10450859','10450879','10450881','10450884','10450908','10510572','10510573','10510574','10510575','10520769','10520771','10520772','10520773','10520774','10520778','10529066','10555406','10555408','10555425','10555427','10555428','10555432','10555434','10555436','10555437','10556631','10556632','10556633','10556634','10556635','10556636','10556637','10556638','10556639','10556640','10556969','10556970','10556971','10556973','10556974','10556975','10556976','10556977','10622014','10622015','10622016','10622017','10633421','10633422','10633423','10633424','10633425','10633426','10667455','10667456','10720428','10720429','10720430','10720431','10733806','10733807','10733808','10733809','10733810','10733811','10733812','10733813','10733814','10733815','10733816','10733817','10733818','10733819','10733820','10733821','10733822','10733823','10733824','10733825','10733826','10733827','10750619','10750620','10754619','10754620','10754622','10763463','10763464','10763465','10763468','10763469','10763470','10763471','10763472','10772225','10772226','10772227','10788515','10794987','101443546','101443685','101443821','101449027','101449045','101449065','101449096','101449117','101449150','101452189','101460274','101460292','101460343','101460368','101460384','101467422','101471482','101480089','101480103','101483272','101483307','101483928','101483932','101483936','101559253','101569721','101569749','101569780','101569823','101569849','101570232','101585502','101585650','101585667','101585683','101585699','101585709','101586247','101590033','101591144','101591155','101593679','101593682','101593722','101593748','101593756','101593781','101593786','101593795','101593803','101594568','101594605','101644898','101644975','101645006','101645009','101645236','101645247','101645349','101645396','101645402','101645407','101645415','101645450','101645455','9818520','10555403','101703255','101703266','101713029','101713036','101713064','101713082','101713088','101713098','101713113','101713137','101713143','101713154','101731264','101731274','101731294','101731510','101731522','101731536','101731549','101731577','101731582','101731590','101745306','101765447')
			)TBL
		GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO		
UNION ALL
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'Vietnam' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
		WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')	
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
		FROM(
			SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
			  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
			FROM REPL_ITEM_LOC REPL
				LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
				LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
				LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
				LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
				LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
				LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
				LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('5806270','9126356','9126363','9126364','9326427','9327072','9409390','9529917','9529920','9624478','9624482','9624484','9702208','9847459','10068734','10073458','10166033','10166045','10166050','10166081','10166088','10166108','10166112','10245435','10245436','10245441','10350119','10350134','10473204','10473206','10473213','10473215','10473224','10473226','10473228','10473230','10688673','10688674','10688675','10688676','10688677','10763474','10763475','10763495','10763500','10763502','10763503','10763504','15000637','15000638','101568160','101568187','101568190','101578139','101578189','101578219','101578729','101578737','101585983','101636658','101647254')
			)TBL
		GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO

};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_lc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "instock_asian.csv: $!";
 
$dbh->disconnect;

}


sub mail1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

#$to = ' cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, jerson.roma@metrogaisano.com, bermon.alcantara@metrogaisano.com, eljie.laquinon@metrogaisano.com, nilynn.yosores@metrogaisano.com, analiza.dano@metrogaisano.com, anafatima.mancho@metrogaisano.com, cindy.yu@metrogaisano.com, emily.silverio@metrogaisano.com, luz.bitang@metrogaisano.com';

#$bcc = 'kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, rex.cabanilla@metrogaisano.com';
$bcc = 'lea.gonzaga@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'In Stock Report, International - Asian Items';

$msgbody_file = 'message.txt';

$attachment_file = 'INSTOCK_REPORT_INTERNATIONAL_v2.xlsx';

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

sub mail2 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

#$to = ' manuel.degamo@metrogaisano.com, ace.olalia@metrogaisano.com, alma.espino@metrogaisano.com, angeli_christi.ladot@metrogaisano.com, angelito.dublin@metrogaisano.com, arlene.yanson@metrogaisano.com, augosto.daria@metrogaisano.com, charm.buenaventura@metrogaisano.com, teena.velasco@metrogaisano.com, cristy.sy@metrogaisano.com, diana.almagro@metrogaisano.com, edgardo.lim@metrogaisano.com, edris.tarrobal@metrogaisano.com, fidela.villamor@metrogaisano.com, genaro.felisilda@metrogaisano.com, genevive.quinones@metrogaisano.com, glenda.navares@metrogaisano.com, joefrey.camu@metrogaisano.com, jonalyn.diaz@metrogaisano.com ';

$bcc = 'lea.gonzaga@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'In Stock Report, International - Asian Items';

$msgbody_file = 'message.txt';

$attachment_file = 'INSTOCK_REPORT_INTERNATIONAL_v2.xlsx';

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

sub mail3 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

#$to = ' josemarie.graciadas@metrogaisano.com, jovany.polancos@metrogaisano.com, judy.gilo@metrogaisano.com, julie.montano@metrogaisano.com, kathlene.procianos@metrogaisano.com, limuel.ulanday@metrogaisano.com, cristina.de_asis@metrogaisano.com, mariajoana.cruz@metrogaisano.com, may.sasedor@metrogaisano.com, michelle.calsada@metrogaisano.com, policarpo.mission@metrogaisano.com, rex.refuerzo@metrogaisano.com, ricky.tulda@metrogaisano.com, ronald.dizon@metrogaisano.com, roselle.agbayani@metrogaisano.com, rowena.tangoan@metrogaisano.com, roy.igot@metrogaisano.com, tessie.cabanero@metrogaisano.com, victoria.ferolino@metrogaisano.com, wendel.gallo@metrogaisano.com, juanjose.sibal@metrogaisano.com, julie.montano@metrogaisano.com ';

$bcc = 'lea.gonzaga@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'In Stock Report, International - Asian Items';

$msgbody_file = 'message.txt';

$attachment_file = 'INSTOCK_REPORT_INTERNATIONAL_v2.xlsx';

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





