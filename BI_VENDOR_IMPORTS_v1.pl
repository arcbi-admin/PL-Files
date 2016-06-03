START:

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
use Date::Calc qw( Today Add_Delta_Days Month_to_Text);

($year,$month,$day) = Today();
$month_to_text = Month_to_Text($month);

$test_query = qq{ SELECT CASE WHEN EXISTS (SELECT *
					FROM ADMIN_ETL_LOG 
					WHERE TO_DATE(LOG_DATE, 'DD-MON-YY') = TO_DATE(SYSDATE,'DD-MON-YY') AND TASK_ID = 'AggDlyStrDept' AND ERR_CODE = 0) THEN 1 ELSE 0 END STATUS 
					FROM DUAL };

$tst_query = $dbh->prepare($test_query);
$tst_query->execute();

while ( my $x =  $tst_query->fetchrow_hashref()){
	$test = $x->{STATUS};
} 

if ($test eq 1){
	 
	 $date = qq{ 
	SELECT WEEK_NUMBER_THIS_YEAR, DATE_KEY, TO_CHAR(DATE_FLD, 'DD Mon YYYY') DATE_FLD, WEEK_ST_DATE_KEY, WEEK_END_DATE_KEY 
	FROM DIM_DATE
	WHERE DATE_FLD = (SELECT AGG_DLY_END_DATE_FLD FROM ADMIN_ETL_DATE_PARAMETER)
	 };

	my $sth_date_1 = $dbh->prepare ($date);
	 $sth_date_1->execute;

	while (my $x = $sth_date_1->fetchrow_hashref()) {
		$wk_st_date_key = $x->{WEEK_ST_DATE_KEY};
		$wk_en_date_key = $x->{DATE_KEY};
		$wk_number = $x->{WEEK_NUMBER_THIS_YEAR};
		$as_of = $x->{DATE_FLD};
	}

	$date_2 = qq{ 
	SELECT DATE_KEY1, TO_CHAR(DATE_FLD1, 'DD Mon YYYY') DATE_FLD1, DATE_KEY_LY1, TO_CHAR(DATE_FLD_LY1, 'DD Mon YYYY') DATE_FLD_LY1, DATE_KEY2, TO_CHAR(DATE_FLD2, 'DD Mon YYYY') DATE_FLD2, DATE_KEY_LY2, TO_CHAR(DATE_FLD_LY2, 'DD Mon YYYY') DATE_FLD_LY2, DATE_KEY3, DATE_FLD3, MONTH_ST_DATE_KEY, MONTH_END_DATE_KEY, QUARTER, YEAR FROM
		(SELECT DATE_KEY AS DATE_KEY1, DATE_FLD AS DATE_FLD1, DATE_KEY_LY AS DATE_KEY_LY1, DATE_FLD_LY AS DATE_FLD_LY1
		FROM DIM_DATE WHERE DATE_KEY = $wk_st_date_key),
		(SELECT DATE_KEY AS DATE_KEY2, DATE_FLD AS DATE_FLD2, DATE_KEY_LY AS DATE_KEY_LY2, DATE_FLD_LY AS DATE_FLD_LY2
		FROM DIM_DATE WHERE DATE_KEY = $wk_en_date_key),
		(SELECT DATE_KEY AS DATE_KEY3, DATE_FLD AS DATE_FLD3, MONTH_ST_DATE_KEY, MONTH_END_DATE_KEY
		FROM DIM_DATE_PRL WHERE TO_CHAR(DATE_FLD, 'DD Mon YYYY') = '$as_of'),
		(SELECT QUARTER, YEAR FROM DIM_DATE_PRL WHERE TO_CHAR(DATE_FLD, 'DD Mon YYYY') = '$as_of')
	 };

	my $sth_date_2 = $dbh->prepare ($date_2);
	 $sth_date_2->execute;
	 
	while (my $x = $sth_date_2->fetchrow_hashref()) {
		$wk_st_date_key_ly = $x->{DATE_KEY_LY1};
		$wk_en_date_key_ly = $x->{DATE_KEY_LY2};
		$wk_st_date_fld = $x->{DATE_FLD1};
		$wk_en_date_fld = $x->{DATE_FLD2};
		$wk_st_date_fld_ly = $x->{DATE_FLD_LY1};
		$wk_en_date_fld_ly = $x->{DATE_FLD_LY2};
		$mo_st_date_key = $x->{MONTH_ST_DATE_KEY};
		$mo_en_date_key = $x->{DATE_KEY3};
		$quarter = $x->{QUARTER};
		$year = $x->{YEAR};
	}

	$date_3 = qq{ 
	SELECT DATE_KEY1, TO_CHAR(DATE_FLD1, 'DD Mon YYYY') DATE_FLD1, DATE_KEY_LY1, TO_CHAR(DATE_FLD_LY1, 'DD Mon YYYY') DATE_FLD_LY1, 
		   DATE_KEY2, TO_CHAR(DATE_FLD2, 'DD Mon YYYY') DATE_FLD2, DATE_KEY_LY2, TO_CHAR(DATE_FLD_LY2, 'DD Mon YYYY') DATE_FLD_LY2,
		   DATE_KEY3, TO_CHAR(DATE_FLD3, 'DD Mon YYYY') DATE_FLD3, DATE_KEY_LY3, TO_CHAR(DATE_FLD_LY3, 'DD Mon YYYY') DATE_FLD_LY3 FROM
		(SELECT DATE_KEY AS DATE_KEY1, DATE_FLD AS DATE_FLD1, DATE_KEY_LY AS DATE_KEY_LY1, DATE_FLD_LY AS DATE_FLD_LY1
		FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_st_date_key),
		(SELECT DATE_KEY AS DATE_KEY2, DATE_FLD AS DATE_FLD2, DATE_KEY_LY AS DATE_KEY_LY2, DATE_FLD_LY AS DATE_FLD_LY2
		FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_en_date_key),
		(SELECT DATE_KEY AS DATE_KEY3, DATE_FLD AS DATE_FLD3, DATE_KEY_LY AS DATE_KEY_LY3, DATE_FLD_LY AS DATE_FLD_LY3
		FROM DIM_DATE_PRL WHERE QUARTER = $quarter AND YEAR = $year AND MONTH_IN_QUARTER = 1 AND DAY_IN_MONTH = 1)
	 };

	my $sth_date_3 = $dbh->prepare ($date_3);
	 $sth_date_3->execute;
	 
	while (my $x = $sth_date_3->fetchrow_hashref()) {
		$mo_st_date_key_ly = $x->{DATE_KEY_LY1};
		$mo_en_date_key_ly = $x->{DATE_KEY_LY2};
		$mo_st_date_fld = $x->{DATE_FLD1};
		$mo_en_date_fld = $x->{DATE_FLD2};
		$mo_st_date_fld_ly = $x->{DATE_FLD_LY1};
		$mo_en_date_fld_ly = $x->{DATE_FLD_LY2};
		$qu_st_date_key = $x->{DATE_KEY3};
		$qu_st_date_key_ly = $x->{DATE_KEY_LY3};
		$qu_st_date_fld = $x->{DATE_FLD3};
		$qu_st_date_fld_ly = $x->{DATE_FLD_LY3};
	}
 
	$workbook = Excel::Writer::XLSX->new("Vendor Imports Sales Performance - Summary (as of $as_of) v1.xlsx");
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
	$headN = $workbook->add_format( border => 1, align => 'center', valign => 'center', size => 11, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	$headD = $workbook->add_format( border => 1, align => 'center', valign => 'center', size => 10, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	$headDPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	$headPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	$headNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3, bg_color => $abo, bold => 1 );
	$headNumber = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 3, bg_color => $abo, bold => 1 );
	$head = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	$subhead = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => '0.0 %', bg_color => $ponkan, bold => 1 );
	$bodyN = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10, bg_color => $sky, num_format => '0.0 %',  bold => 1);
	$bodyPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => '0.0 %',  bold => 1);
	$bodyNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 3,  bold => 1);
	$body = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => '0.0 %',  bold => 1);
	$subt = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => '0.0 %');
	$down = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => '0.0 %', bg_color => $pula );

	printf "Test ETL Status = ". $test ." \nArc BI Sales Performance Part 1\nGenerating Data from Source \n";
	
	# &generate_data;
	# &insert_data;
	
	&new_sheet_2($sheet = "Department");			
	&call_div;
	
	&new_sheet_2($sheet = "Country");			
	&call_div_country;
	
	&new_sheet_2($sheet = "Vendor");			
	&call_div_vendor;
		
	$workbook->close();
	
	$tst_query->finish();
	$dbh->disconnect; 
	
	exit;
	
}

elsif ($test eq 0){
	print "Test Status = ". $test ." \nETL still running\n";
	
	$tst_query->finish();
	
	sleep(300);
	
	goto START;
}
 
#================================= FUNCTIONS ==================================#

sub call_div {

$a = 10, $e = 10, $counter = 0;

$imports_mtd = 0, $total_mtd = 0, $imports_ly_mtd = 0, $total_ly_mtd = 0, $imports_qtd = 0, $total_qtd = 0, $imports_ly_qtd = 0, $total_ly_qtd = 0;
$total_imports_mtd = 0, $total_total_mtd = 0, $total_imports_ly_mtd = 0, $total_total_ly_mtd = 0, $total_imports_qtd = 0, $total_total_qtd = 0, $total_imports_ly_qtd = 0, $total_total_ly_qtd = 0;

$type_test = 0;

$worksheet->write($a-10, 2, "Supermarket Imports Performance", $bold1);
#$worksheet->write($a-9, 2, "WTD: $wk_st_date_fld - $mo_en_date_fld vs $wk_st_date_fld_ly - $wk_en_date_fld_ly");
$worksheet->write($a-8, 2, "MTD: $mo_st_date_fld - $mo_en_date_fld vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-7, 2, "QTD: $qu_st_date_fld - $mo_en_date_fld vs $qu_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-6, 2, "As of $as_of");

##========================= COMP STORES ===========================##

&heading_2;
&heading;
&query_dept($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 1, $matured_flg2 = 1, $loc_desc = "COMP STORES");

##========================= ALL STORES ===========================##

$a += 7;

$imports_mtd = 0, $total_mtd = 0, $imports_ly_mtd = 0, $total_ly_mtd = 0, $imports_qtd = 0, $total_qtd = 0, $imports_ly_qtd = 0, $total_ly_qtd = 0;
$total_imports_mtd = 0, $total_total_mtd = 0, $total_imports_ly_mtd = 0, $total_total_ly_mtd = 0, $total_imports_qtd = 0, $total_total_qtd = 0, $total_imports_ly_qtd = 0, $total_total_ly_qtd = 0;

$type_test = 0;

&heading_2;
&heading;
&query_dept($new_flg1 = 0, $new_flg2 = 1, $matured_flg1 = 0, $matured_flg2 = 1, $loc_desc = "ALL STORES");

##========================= BY STORE ===========================##

foreach my $i ( '2001', '4002', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', '2013', '2223', '3001', '3002', '3003', '3004', '3005', '3006', '3007', '3009', '3012', '4003', '4004', '6001', '6002', '6003', '6004', '6005', '6009', '6010', '6012' ){ 
# foreach my $i ( '2001', '4002' ){ 
	$a += 7;
	
	$imports_mtd = 0, $total_mtd = 0, $imports_ly_mtd = 0, $total_ly_mtd = 0, $imports_qtd = 0, $total_qtd = 0, $imports_ly_qtd = 0, $total_ly_qtd = 0;
	$total_imports_mtd = 0, $total_total_mtd = 0, $total_imports_ly_mtd = 0, $total_total_ly_mtd = 0, $total_imports_qtd = 0, $total_total_qtd = 0, $total_imports_ly_qtd = 0, $total_total_ly_qtd = 0;
	
	&heading_2;
	&heading;
	&query_dept_store($store = $i);

}

}

sub call_div_country {

$a = 10, $e = 10, $counter = 0;

$type_test = 0;

$worksheet->write($a-10, 2, "Supermarket Imports Performance", $bold1);
#$worksheet->write($a-9, 2, "WTD: $wk_st_date_fld - $mo_en_date_fld vs $wk_st_date_fld_ly - $wk_en_date_fld_ly");
$worksheet->write($a-8, 2, "MTD: $mo_st_date_fld - $mo_en_date_fld vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-7, 2, "QTD: $qu_st_date_fld - $mo_en_date_fld vs $qu_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-6, 2, "As of $as_of");

##========================= ALL STORES ===========================##

$total_d_imports_mtd = 0, $total_d_imports_ly_mtd = 0, $total_l_imports_mtd = 0, $total_l_imports_ly_mtd = 0, $total_imports_mtd = 0, $total_imports_ly_mtd = 0, $total_d_imports_qtd = 0, $total_d_imports_ly_qtd = 0, $total_l_imports_qtd = 0, $total_l_imports_ly_qtd = 0, $total_imports_qtd = 0, $total_imports_ly_qtd = 0;

$type_test = 0;

&heading_2;
&heading_country;
&query_dept_country($loc_desc = "ALL STORES");


}

sub call_div_vendor {

$a = 10, $e = 10, $counter = 0;

$type_test = 0;

$worksheet->write($a-10, 2, "Supermarket Imports Performance", $bold1);
#$worksheet->write($a-9, 2, "WTD: $wk_st_date_fld - $mo_en_date_fld vs $wk_st_date_fld_ly - $wk_en_date_fld_ly");
$worksheet->write($a-8, 2, "MTD: $mo_st_date_fld - $mo_en_date_fld vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-7, 2, "QTD: $qu_st_date_fld - $mo_en_date_fld vs $qu_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-6, 2, "As of $as_of");

##========================= ALL STORES ===========================##

$total_mtd = 0, $total_ly_mtd = 0, $total_qtd = 0, $total_ly_qtd = 0;
$total_total_mtd = 0, $total_total_ly_mtd = 0, $total_total_qtd = 0, $total_total_ly_qtd = 0;

$type_test = 0;

&heading_2;
&heading_ven;
&query_dept_vendor($loc_desc = "ALL STORES");


}


sub new_sheet_2{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(90);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
$worksheet->set_print_scale( 100 );
$worksheet->set_margins( 0.001 );
$worksheet->conditional_formatting( 'F9:AK6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );
$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 23 );

$worksheet->set_column( 7, 20, 9 );

}


sub heading {

$worksheet->write($a-3, 3, "in 000's", $script);

$worksheet->merge_range( $a-3, 7, $a-3, 13, 'MTD', $subhead );
$worksheet->merge_range( $a-3, 14, $a-3, 20, 'QTD', $subhead );

foreach my $i ( 7, 10, 14, 17 ) {
	$worksheet->write($a-1, $i, "Imports", $subhead);
	$worksheet->write($a-1, $i+1, "Dept Sales", $subhead);
	$worksheet->write($a-1, $i+2, "% to Sales", $subhead);
	
	if ($i eq 7 or $i eq 14){
		$worksheet->merge_range( $a-2, $i, $a-2, $i+2, 'TY', $subhead );
		$worksheet->merge_range( $a-2, $i+3, $a-2, $i+5, 'LY', $subhead );
	}
	
	elsif ($i eq 10 or $i eq 17){
		$worksheet->merge_range( $a-2, $i+3, $a-1, $i+3, 'Growth', $subhead );
	}
}

}

sub heading_2 {

$loc = $a-4;

$worksheet->merge_range( $a-2, 2, $a-1, 2, 'Type', $subhead );
$worksheet->merge_range( $a-2, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a-1, 4, 'Type', $subhead );
$worksheet->merge_range( $a-2, 5, $a-1, 5, 'Code', $subhead );
$worksheet->merge_range( $a-2, 6, $a-1, 6, 'Desc', $subhead );

}

sub heading_ven {

$worksheet->write($a-3, 3, "in 000's", $script);

$worksheet->merge_range( $a-2, 7, $a-2, 9, 'MTD', $subhead );
$worksheet->merge_range( $a-2, 10, $a-2, 12, 'QTD', $subhead );

foreach my $i ( 7, 10 ) {
	$worksheet->write($a-1, $i, "TY", $subhead);
	$worksheet->write($a-1, $i+1, "LY", $subhead);
	$worksheet->write($a-1, $i+2, "% Growth", $subhead);	
}

$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, 8 );
$worksheet->set_column( 6, 6, 35 );

$worksheet->set_column( 7, 12, 9 );

}

sub heading_country {

$worksheet->write($a-3, 3, "in 000's", $script);

$worksheet->merge_range( $a-3, 7, $a-3, 15, 'MTD', $subhead );
$worksheet->merge_range( $a-3, 16, $a-3, 24, 'QTD', $subhead );

foreach my $i ( 7, 10, 13, 16, 19, 22 ) {
	$worksheet->write($a-1, $i, "TY", $subhead);
	$worksheet->write($a-1, $i+1, "LY", $subhead);
	$worksheet->write($a-1, $i+2, "Growth", $subhead);
	
	if ($i eq 7 or $i eq 16){
		$worksheet->merge_range( $a-2, $i, $a-2, $i+2, 'Direct', $subhead );
		$worksheet->merge_range( $a-2, $i+3, $a-2, $i+5, 'Local', $subhead );
		$worksheet->merge_range( $a-2, $i+6, $a-2, $i+8, 'Total', $subhead );
	}

}

$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 20 );

$worksheet->set_column( 7, 12, 7 );

}

# sheet 3
sub query_dept {
	
$sls1 = $dbh->prepare (qq{SELECT tot.merch_group_code, tot.merch_group_desc, 
						imports_mtd, total_mtd, case when total_mtd = 0 then null else round(imports_mtd/total_mtd,3) end as cont_mtd, 
						imports_ly_mtd, total_ly_mtd, case when total_ly_mtd = 0 then null else round(imports_ly_mtd/total_ly_mtd,3) end as cont_ly_mtd,
						imports_qtd, total_qtd, case when total_qtd = 0 then null else round(imports_qtd/total_qtd,3) end as cont_qtd,
						imports_ly_qtd, total_ly_qtd, case when total_ly_qtd = 0 then null else round(imports_ly_qtd/total_ly_qtd,3) end as cont_ly_qtd 
						FROM
							(SELECT merch_group_code, merch_group_desc, SUM(total_retail_mtd)/1000 total_mtd, SUM(total_retail_ly_mtd)/1000 total_ly_mtd, SUM(total_retail_qtd)/1000 total_qtd, SUM(total_retail_ly_qtd)/1000 total_ly_qtd 
							FROM metro_it_vendor_imports4
							WHERE ((new_flg = '$new_flg1' or new_flg = '$new_flg2') and (matured_flg = '$matured_flg1' or matured_flg = '$matured_flg2'))
							GROUP BY merch_group_code, merch_group_desc)TOT
							LEFT JOIN
							(SELECT merch_group_code, merch_group_desc, SUM(total_retail_mtd)/1000 imports_mtd, SUM(total_retail_ly_mtd)/1000 imports_ly_mtd, SUM(total_retail_qtd)/1000 imports_qtd, SUM(total_retail_ly_qtd)/1000 imports_ly_qtd
							FROM metro_it_vendor_imports4
							WHERE ((new_flg = '$new_flg1' OR new_flg = '$new_flg2') AND (matured_flg = '$matured_flg1' OR matured_flg = '$matured_flg2')) AND country IS NOT NULL
							GROUP BY merch_group_code, merch_group_desc)imp
							ON tot.merch_group_code = imp.merch_group_code
						ORDER BY 1
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{MERCH_GROUP_CODE};
	$merch_group_desc = $s->{MERCH_GROUP_DESC};
	
	$sls2 = $dbh->prepare (qq{SELECT tot.group_code, tot.group_desc,  
								imports_mtd, total_mtd, case when total_mtd = 0 then null else round(imports_mtd/total_mtd,3) end as cont_mtd, 
								imports_ly_mtd, total_ly_mtd, case when total_ly_mtd = 0 then null else round(imports_ly_mtd/total_ly_mtd,3) end as cont_ly_mtd,
								imports_qtd, total_qtd, case when total_qtd = 0 then null else round(imports_qtd/total_qtd,3) end as cont_qtd,
								imports_ly_qtd, total_ly_qtd, case when total_ly_qtd = 0 then null else round(imports_ly_qtd/total_ly_qtd,3) end as cont_ly_qtd
								FROM
									(SELECT group_code, group_desc, SUM(total_retail_mtd)/1000 total_mtd, SUM(total_retail_ly_mtd)/1000 total_ly_mtd, SUM(total_retail_qtd)/1000 total_qtd, SUM(total_retail_ly_qtd)/1000 total_ly_qtd 
									FROM metro_it_vendor_imports4
									WHERE merch_group_code = '$merch_group_code' and ((new_flg = '$new_flg1' or new_flg = '$new_flg2') and (matured_flg = '$matured_flg1' or matured_flg = '$matured_flg2'))
									GROUP BY group_code, group_desc)TOT
									LEFT JOIN
									(SELECT group_code, group_desc, SUM(total_retail_mtd)/1000 imports_mtd, SUM(total_retail_ly_mtd)/1000 imports_ly_mtd, SUM(total_retail_qtd)/1000 imports_qtd, SUM(total_retail_ly_qtd)/1000 imports_ly_qtd
									FROM metro_it_vendor_imports4
									WHERE merch_group_code = '$merch_group_code' and ((new_flg = '$new_flg1' OR new_flg = '$new_flg2') AND (matured_flg = '$matured_flg1' OR matured_flg = '$matured_flg2')) AND country IS NOT NULL
									GROUP BY group_code, group_desc)imp
									ON tot.group_code = imp.group_code
								ORDER BY 1		
							 });	
	$sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$group_code = $s->{GROUP_CODE};
		$group_desc = $s->{GROUP_DESC};
				
		$sls3 = $dbh->prepare (qq{SELECT tot.division, tot.division_desc,  
									imports_mtd, total_mtd, case when total_mtd = 0 then null else round(imports_mtd/total_mtd,3) end as cont_mtd, 
									imports_ly_mtd, total_ly_mtd, case when total_ly_mtd = 0 then null else round(imports_ly_mtd/total_ly_mtd,3) end as cont_ly_mtd,
									imports_qtd, total_qtd, case when total_qtd = 0 then null else round(imports_qtd/total_qtd,3) end as cont_qtd,
									imports_ly_qtd, total_ly_qtd, case when total_ly_qtd = 0 then null else round(imports_ly_qtd/total_ly_qtd,3) end as cont_ly_qtd 
									FROM
										(SELECT division, division_desc, SUM(total_retail_mtd)/1000 total_mtd, SUM(total_retail_ly_mtd)/1000 total_ly_mtd, SUM(total_retail_qtd)/1000 total_qtd, SUM(total_retail_ly_qtd)/1000 total_ly_qtd 
										FROM metro_it_vendor_imports4
										WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and ((new_flg = '$new_flg1' or new_flg = '$new_flg2') and (matured_flg = '$matured_flg1' or matured_flg = '$matured_flg2'))
										GROUP BY division, division_desc)TOT
										LEFT JOIN
										(SELECT division, division_desc, SUM(total_retail_mtd)/1000 imports_mtd, SUM(total_retail_ly_mtd)/1000 imports_ly_mtd, SUM(total_retail_qtd)/1000 imports_qtd, SUM(total_retail_ly_qtd)/1000 imports_ly_qtd
										FROM metro_it_vendor_imports4
										WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and ((new_flg = '$new_flg1' OR new_flg = '$new_flg2') AND (matured_flg = '$matured_flg1' OR matured_flg = '$matured_flg2')) AND country IS NOT NULL
										GROUP BY division, division_desc)imp
										ON tot.division = imp.division
									ORDER BY 1	
									});
		$sls3->execute();
		
		$grp_counter = $a;
		while(my $s = $sls3->fetchrow_hashref()){
			$division = $s->{DIVISION};
			$division_desc = $s->{DIVISION_DESC};
			
			$sls4 = $dbh->prepare (qq{SELECT tot.group_no department_code, tot.group_name department_desc,  
										imports_mtd, total_mtd, case when total_mtd = 0 then null else round(imports_mtd/total_mtd,3) end as cont_mtd, 
										imports_ly_mtd, total_ly_mtd, case when total_ly_mtd = 0 then null else round(imports_ly_mtd/total_ly_mtd,3) end as cont_ly_mtd,
										imports_qtd, total_qtd, case when total_qtd = 0 then null else round(imports_qtd/total_qtd,3) end as cont_qtd,
										imports_ly_qtd, total_ly_qtd, case when total_ly_qtd = 0 then null else round(imports_ly_qtd/total_ly_qtd,3) end as cont_ly_qtd 
										FROM
											(SELECT group_no, group_name, SUM(total_retail_mtd)/1000 total_mtd, SUM(total_retail_ly_mtd)/1000 total_ly_mtd, SUM(total_retail_qtd)/1000 total_qtd, SUM(total_retail_ly_qtd)/1000 total_ly_qtd 
											FROM metro_it_vendor_imports4
											WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and division = '$division' and ((new_flg = '$new_flg1' or new_flg = '$new_flg2') and (matured_flg = '$matured_flg1' or matured_flg = '$matured_flg2'))
											GROUP BY group_no, group_name)TOT
											LEFT JOIN
											(SELECT group_no, group_name, SUM(total_retail_mtd)/1000 imports_mtd, SUM(total_retail_ly_mtd)/1000 imports_ly_mtd, SUM(total_retail_qtd)/1000 imports_qtd, SUM(total_retail_ly_qtd)/1000 imports_ly_qtd
											FROM metro_it_vendor_imports4
											WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and division = '$division' and ((new_flg = '$new_flg1' or new_flg = '$new_flg2') and (matured_flg = '$matured_flg1' or matured_flg = '$matured_flg2')) AND country IS NOT NULL
											GROUP BY group_no, group_name)imp
											ON tot.group_no = imp.group_no
										ORDER BY 1		
										 });
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){
				
				$worksheet->write($a,5, $s->{DEPARTMENT_CODE},$desc);
				$worksheet->write($a,6, $s->{DEPARTMENT_DESC},$desc);
				
				$worksheet->write($a,7, $s->{IMPORTS_MTD},$border1);
				$worksheet->write($a,8, $s->{TOTAL_MTD},$border1);
				$worksheet->write($a,9, $s->{CONT_MTD},$subt);
				
				$worksheet->write($a,10, $s->{IMPORTS_LY_MTD},$border1);					
				$worksheet->write($a,11, $s->{TOTAL_LY_MTD},$border1);
				$worksheet->write($a,12, $s->{CONT_LY_MTD},$subt);
				
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{IMPORTS_QTD},$border1);						
				$worksheet->write($a,15, $s->{TOTAL_QTD},$border1);
				$worksheet->write($a,16, $s->{CONT_QTD},$subt);
				
				$worksheet->write($a,17, $s->{IMPORTS_LY_QTD},$border1);						
				$worksheet->write($a,18, $s->{TOTAL_LY_QTD},$border1);
				$worksheet->write($a,19, $s->{CONT_LY_QTD},$subt);
				
				$worksheet->write($a,20, "",$subt);
				
				$a++;
				$counter++;
		
			}			
			
			$worksheet->write($a,7, $s->{IMPORTS_MTD},$bodyNum);
			$worksheet->write($a,8, $s->{TOTAL_MTD},$bodyNum);
			$worksheet->write($a,9, $s->{CONT_MTD},$bodyPct);
			
			$worksheet->write($a,10, $s->{IMPORTS_LY_MTD},$bodyNum);					
			$worksheet->write($a,11, $s->{TOTAL_LY_MTD},$bodyNum);
			$worksheet->write($a,12, $s->{CONT_LY_MTD},$bodyPct);
			
			$worksheet->write($a,13, "",$bodyPct);
			
			$worksheet->write($a,14, $s->{IMPORTS_QTD},$bodyNum);						
			$worksheet->write($a,15, $s->{TOTAL_QTD},$bodyNum);
			$worksheet->write($a,16, $s->{CONT_QTD},$bodyPct);
			
			$worksheet->write($a,17, $s->{IMPORTS_LY_QTD},$bodyNum);						
			$worksheet->write($a,18, $s->{TOTAL_LY_QTD},$bodyNum);
			$worksheet->write($a,19, $s->{CONT_LY_QTD},$bodyPct);
			
			$worksheet->write($a,20, "",$bodyPct);

			$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
			
			$counter = 0; #RESET dept_counter
			$a++; #INCREMENT VARIABLE a
		}
		
		$worksheet->write($a,7, $s->{IMPORTS_MTD},$bodyNum);
		$worksheet->write($a,8, $s->{TOTAL_MTD},$bodyNum);
		$worksheet->write($a,9, $s->{CONT_MTD},$bodyPct);
		
		$worksheet->write($a,10, $s->{IMPORTS_LY_MTD},$bodyNum);					
		$worksheet->write($a,11, $s->{TOTAL_LY_MTD},$bodyNum);
		$worksheet->write($a,12, $s->{CONT_LY_MTD},$bodyPct);
		
		$worksheet->write($a,13, "",$bodyPct);
		
		$worksheet->write($a,14, $s->{IMPORTS_QTD},$bodyNum);						
		$worksheet->write($a,15, $s->{TOTAL_QTD},$bodyNum);
		$worksheet->write($a,16, $s->{CONT_QTD},$bodyPct);
		
		$worksheet->write($a,17, $s->{IMPORTS_LY_QTD},$bodyNum);						
		$worksheet->write($a,18, $s->{TOTAL_LY_QTD},$bodyNum);
		$worksheet->write($a,19, $s->{CONT_LY_QTD},$bodyPct);
		
		$worksheet->write($a,20, "",$bodyPct);

		$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

		$a++; #INCREMENT VARIABLE a
	}
	
	$total_imports_mtd += $s->{IMPORTS_MTD};
	$total_total_mtd += $s->{TOTAL_MTD};
			
	$total_imports_ly_mtd += $s->{IMPORTS_LY_MTD};
	$total_total_ly_mtd += $s->{TOTAL_LY_MTD};
			
	$total_imports_qtd += $s->{IMPORTS_QTD};
	$total_total_qtd += $s->{TOTAL_QTD};
			
	$total_imports_ly_qtd += $s->{IMPORTS_LY_QTD};
	$total_total_ly_qtd += $s->{TOTAL_LY_QTD};
	
	if ($merch_group_code eq 'DS'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
	}
	
	elsif($merch_group_code eq 'SU'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
	}
	
	elsif($merch_group_code eq 'Z_OT'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Others', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'OTHERS', $border2 );
	}
		
	$worksheet->write($a,7, $s->{IMPORTS_MTD},$headNumber);
	$worksheet->write($a,8, $s->{TOTAL_MTD},$headNumber);
	$worksheet->write($a,9, $s->{CONT_MTD},$headPct);
	
	$worksheet->write($a,10, $s->{IMPORTS_LY_MTD},$headNumber);					
	$worksheet->write($a,11, $s->{TOTAL_LY_MTD},$headNumber);
	$worksheet->write($a,12, $s->{CONT_LY_MTD},$headPct);
	
	$worksheet->write($a,13, "",$headPct);
	
	$worksheet->write($a,14, $s->{IMPORTS_QTD},$headNumber);						
	$worksheet->write($a,15, $s->{TOTAL_QTD},$headNumber);
	$worksheet->write($a,16, $s->{CONT_QTD},$headPct);
	
	$worksheet->write($a,17, $s->{IMPORTS_LY_QTD},$headNumber);						
	$worksheet->write($a,18, $s->{TOTAL_LY_QTD},$headNumber);
	$worksheet->write($a,19, $s->{CONT_LY_QTD},$headPct);
	
	$worksheet->write($a,20, "",$headPct);
	
	$a++; #INCREMENT VARIABLE a
}
	
	$worksheet->write($a,7, $total_imports_mtd, $headNumber);
	$worksheet->write($a,8, $total_total_mtd, $headNumber);
		if ($total_total_mtd eq 0) { $worksheet->write($a,9, "", $headPct); }
		else { $worksheet->write($a,9, $total_imports_mtd/$total_total_mtd, $headPct); }
	
	$worksheet->write($a,10, $total_imports_ly_mtd, $headNumber);					
	$worksheet->write($a,11, $total_total_ly_mtd, $headNumber);
		if ($total_total_ly_mtd eq 0) { $worksheet->write($a,12, "", $headPct); }
		else { $worksheet->write($a,12, $total_imports_ly_mtd/$total_total_ly_mtd, $headPct); }
		
	$worksheet->write($a,13, "",$headPct);
	
	$worksheet->write($a,14, $total_imports_qtd, $headNumber);						
	$worksheet->write($a,15, $total_total_qtd, $headNumber);
		if ($total_total_qtd eq 0) { $worksheet->write($a,16, "", $headPct); }
		else { $worksheet->write($a,16, $total_imports_qtd/$total_total_qtd, $headPct); }
	
	$worksheet->write($a,17, $total_imports_ly_qtd, $headNumber);						
	$worksheet->write($a,18, $total_total_ly_qtd, $headNumber);
		if ($total_total_ly_qtd eq 0) { $worksheet->write($a,19, "", $headPct); }
		else { $worksheet->write($a,19, $total_imports_ly_qtd/$total_total_ly_qtd, $headPct); }
	
	$worksheet->write($a,20, "",$headPct);
	
$worksheet->write($loc, 2, $loc_desc, $bold);
$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}

sub query_dept_store {
			
$sls1 = $dbh->prepare (qq{SELECT tot.location store_code, tot.location store_description, tot.merch_group_code, 
						imports_mtd, total_mtd, case when total_mtd = 0 then null else round(imports_mtd/total_mtd,3) end as cont_mtd, 
						imports_ly_mtd, total_ly_mtd, case when total_ly_mtd = 0 then null else round(imports_ly_mtd/total_ly_mtd,3) end as cont_ly_mtd,
						imports_qtd, total_qtd, case when total_qtd = 0 then null else round(imports_qtd/total_qtd,3) end as cont_qtd,
						imports_ly_qtd, total_ly_qtd, case when total_ly_qtd = 0 then null else round(imports_ly_qtd/total_ly_qtd,3) end as cont_ly_qtd
						FROM
							(SELECT location, merch_group_code, SUM(total_retail_mtd)/1000 total_mtd, SUM(total_retail_ly_mtd)/1000 total_ly_mtd, SUM(total_retail_qtd)/1000 total_qtd, SUM(total_retail_ly_qtd)/1000 total_ly_qtd 
							FROM metro_it_vendor_imports4
							WHERE location = '$store'
							GROUP BY location, merch_group_code)TOT
							LEFT JOIN
							(SELECT location, merch_group_code, SUM(total_retail_mtd)/1000 imports_mtd, SUM(total_retail_ly_mtd)/1000 imports_ly_mtd, SUM(total_retail_qtd)/1000 imports_qtd, SUM(total_retail_ly_qtd)/1000 imports_ly_qtd
							FROM metro_it_vendor_imports4
							WHERE country IS NOT NULL and location = '$store'
							GROUP BY location, merch_group_code)imp
							ON tot.location = imp.location and tot.merch_group_code = imp.merch_group_code
						ORDER BY 1, 3
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{MERCH_GROUP_CODE};
	$loc_code = $s->{STORE_CODE};
	$loc_desc = $s->{STORE_DESCRIPTION};
	
	$sls2 = $dbh->prepare (qq{SELECT tot.group_code, tot.group_desc,  
							imports_mtd, total_mtd, case when total_mtd = 0 then null else round(imports_mtd/total_mtd,3) end as cont_mtd, 
							imports_ly_mtd, total_ly_mtd, case when total_ly_mtd = 0 then null else round(imports_ly_mtd/total_ly_mtd,3) end as cont_ly_mtd,
							imports_qtd, total_qtd, case when total_qtd = 0 then null else round(imports_qtd/total_qtd,3) end as cont_qtd,
							imports_ly_qtd, total_ly_qtd, case when total_ly_qtd = 0 then null else round(imports_ly_qtd/total_ly_qtd,3) end as cont_ly_qtd 
								FROM
									(SELECT group_code, group_desc, SUM(total_retail_mtd)/1000 total_mtd, SUM(total_retail_ly_mtd)/1000 total_ly_mtd, SUM(total_retail_qtd)/1000 total_qtd, SUM(total_retail_ly_qtd)/1000 total_ly_qtd 
									FROM metro_it_vendor_imports4
									WHERE merch_group_code = '$merch_group_code' and location = '$store'
									GROUP BY group_code, group_desc)TOT
									LEFT JOIN
									(SELECT group_code, group_desc, SUM(total_retail_mtd)/1000 imports_mtd, SUM(total_retail_ly_mtd)/1000 imports_ly_mtd, SUM(total_retail_qtd)/1000 imports_qtd, SUM(total_retail_ly_qtd)/1000 imports_ly_qtd
									FROM metro_it_vendor_imports4
									WHERE merch_group_code = '$merch_group_code' and location = '$store' and country IS NOT NULL
									GROUP BY group_code, group_desc)imp
									ON tot.group_code = imp.group_code
								ORDER BY 1	
							 });	
	$sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$group_code = $s->{GROUP_CODE};
		$group_desc = $s->{GROUP_DESC};
				
		$sls3 = $dbh->prepare (qq{SELECT tot.division, tot.division_desc,  
								imports_mtd, total_mtd, case when total_mtd = 0 then null else round(imports_mtd/total_mtd,3) end as cont_mtd, 
								imports_ly_mtd, total_ly_mtd, case when total_ly_mtd = 0 then null else round(imports_ly_mtd/total_ly_mtd,3) end as cont_ly_mtd,
								imports_qtd, total_qtd, case when total_qtd = 0 then null else round(imports_qtd/total_qtd,3) end as cont_qtd,
								imports_ly_qtd, total_ly_qtd, case when total_ly_qtd = 0 then null else round(imports_ly_qtd/total_ly_qtd,3) end as cont_ly_qtd
									FROM
										(SELECT division, division_desc, SUM(total_retail_mtd)/1000 total_mtd, SUM(total_retail_ly_mtd)/1000 total_ly_mtd, SUM(total_retail_qtd)/1000 total_qtd, SUM(total_retail_ly_qtd)/1000 total_ly_qtd 
										FROM metro_it_vendor_imports4
										WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and location = '$store'
										GROUP BY division, division_desc)TOT
										LEFT JOIN
										(SELECT division, division_desc, SUM(total_retail_mtd)/1000 imports_mtd, SUM(total_retail_ly_mtd)/1000 imports_ly_mtd, SUM(total_retail_qtd)/1000 imports_qtd, SUM(total_retail_ly_qtd)/1000 imports_ly_qtd
										FROM metro_it_vendor_imports4
										WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and location = '$store' AND country IS NOT NULL
										GROUP BY division, division_desc)imp
										ON tot.division = imp.division
									ORDER BY 1	
									});
		$sls3->execute();
		
		$grp_counter = $a;
		while(my $s = $sls3->fetchrow_hashref()){
			$division = $s->{DIVISION};
			$division_desc = $s->{DIVISION_DESC};
			
			$sls4 = $dbh->prepare (qq{SELECT tot.group_no department_code, tot.group_name department_desc,  
										imports_mtd, total_mtd, case when total_mtd = 0 then null else round(imports_mtd/total_mtd,3) end as cont_mtd, 
										imports_ly_mtd, total_ly_mtd, case when total_ly_mtd = 0 then null else round(imports_ly_mtd/total_ly_mtd,3) end as cont_ly_mtd,
										imports_qtd, total_qtd, case when total_qtd = 0 then null else round(imports_qtd/total_qtd,3) end as cont_qtd,
										imports_ly_qtd, total_ly_qtd, case when total_ly_qtd = 0 then null else round(imports_ly_qtd/total_ly_qtd,3) end as cont_ly_qtd FROM
											(SELECT group_no, group_name, SUM(total_retail_mtd)/1000 total_mtd, SUM(total_retail_ly_mtd)/1000 total_ly_mtd, SUM(total_retail_qtd)/1000 total_qtd, SUM(total_retail_ly_qtd)/1000 total_ly_qtd 
											FROM metro_it_vendor_imports4
											WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and division = '$division' and location = '$store'
											GROUP BY group_no, group_name)TOT
											LEFT JOIN
											(SELECT group_no, group_name, SUM(total_retail_mtd)/1000 imports_mtd, SUM(total_retail_ly_mtd)/1000 imports_ly_mtd, SUM(total_retail_qtd)/1000 imports_qtd, SUM(total_retail_ly_qtd)/1000 imports_ly_qtd
											FROM metro_it_vendor_imports4
											WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and division = '$division' and location = '$store' AND country IS NOT NULL
											GROUP BY group_no, group_name)imp
											ON tot.group_no = imp.group_no
										ORDER BY 1	
										 });
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){
			
				$worksheet->write($a,5, $s->{DEPARTMENT_CODE},$desc);
				$worksheet->write($a,6, $s->{DEPARTMENT_DESC},$desc);
				
				$worksheet->write($a,7, $s->{IMPORTS_MTD},$border1);
				$worksheet->write($a,8, $s->{TOTAL_MTD},$border1);
				$worksheet->write($a,9, $s->{CONT_MTD},$subt);
				
				$worksheet->write($a,10, $s->{IMPORTS_LY_MTD},$border1);					
				$worksheet->write($a,11, $s->{TOTAL_LY_MTD},$border1);
				$worksheet->write($a,12, $s->{CONT_LY_MTD},$subt);
				
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{IMPORTS_QTD},$border1);						
				$worksheet->write($a,15, $s->{TOTAL_QTD},$border1);
				$worksheet->write($a,16, $s->{CONT_QTD},$subt);
				
				$worksheet->write($a,17, $s->{IMPORTS_LY_QTD},$border1);						
				$worksheet->write($a,18, $s->{TOTAL_LY_QTD},$border1);
				$worksheet->write($a,19, $s->{CONT_LY_QTD},$subt);
				
				$worksheet->write($a,20, "",$subt);
								
				$a++;
				$counter++;
		
			}
			
			$worksheet->write($a,7, $s->{IMPORTS_MTD},$bodyNum);
			$worksheet->write($a,8, $s->{TOTAL_MTD},$bodyNum);
			$worksheet->write($a,9, $s->{CONT_MTD},$bodyPct);
			
			$worksheet->write($a,10, $s->{IMPORTS_LY_MTD},$bodyNum);					
			$worksheet->write($a,11, $s->{TOTAL_LY_MTD},$bodyNum);
			$worksheet->write($a,12, $s->{CONT_LY_MTD},$bodyPct);
			
			$worksheet->write($a,13, "",$bodyPct);
			
			$worksheet->write($a,14, $s->{IMPORTS_QTD},$bodyNum);						
			$worksheet->write($a,15, $s->{TOTAL_QTD},$bodyNum);
			$worksheet->write($a,16, $s->{CONT_QTD},$bodyPct);
			
			$worksheet->write($a,17, $s->{IMPORTS_LY_QTD},$bodyNum);						
			$worksheet->write($a,18, $s->{TOTAL_LY_QTD},$bodyNum);
			$worksheet->write($a,19, $s->{CONT_LY_QTD},$bodyPct);
			
			$worksheet->write($a,13, "",$bodyPct);

			$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
			
			$counter = 0; #RESET dept_counter	
			$a++; #INCREMENT VARIABLE a
		}
		
		$worksheet->write($a,7, $s->{IMPORTS_MTD},$bodyNum);
		$worksheet->write($a,8, $s->{TOTAL_MTD},$bodyNum);
		$worksheet->write($a,9, $s->{CONT_MTD},$bodyPct);
		
		$worksheet->write($a,10, $s->{IMPORTS_LY_MTD},$bodyNum);					
		$worksheet->write($a,11, $s->{TOTAL_LY_MTD},$bodyNum);
		$worksheet->write($a,12, $s->{CONT_LY_MTD},$bodyPct);
		
		$worksheet->write($a,13, "",$bodyPct);
		
		$worksheet->write($a,14, $s->{IMPORTS_QTD},$bodyNum);						
		$worksheet->write($a,15, $s->{TOTAL_QTD},$bodyNum);
		$worksheet->write($a,16, $s->{CONT_QTD},$bodyPct);
		
		$worksheet->write($a,17, $s->{IMPORTS_LY_QTD},$bodyNum);						
		$worksheet->write($a,18, $s->{TOTAL_LY_QTD},$bodyNum);
		$worksheet->write($a,19, $s->{CONT_LY_QTD},$bodyPct);
		
		$worksheet->write($a,20, "",$bodyPct);

		$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

		$a++; #INCREMENT VARIABLE a
	}
	
	$total_imports_mtd += $s->{IMPORTS_MTD};
	$total_total_mtd += $s->{TOTAL_MTD};
			
	$total_imports_ly_mtd += $s->{IMPORTS_LY_MTD};
	$total_total_ly_mtd += $s->{TOTAL_LY_MTD};
			
	$total_imports_qtd += $s->{IMPORTS_QTD};
	$total_total_qtd += $s->{TOTAL_QTD};
			
	$total_imports_ly_qtd += $s->{IMPORTS_LY_QTD};
	$total_total_ly_qtd += $s->{TOTAL_LY_QTD};
	
	if ($merch_group_code eq 'DS'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 ); }
		
	elsif($merch_group_code eq 'SU'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 ); }
		
	elsif($merch_group_code eq 'Z_OT'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Others', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'OTHERS', $border2 ); }
	
	$worksheet->write($a,7, $s->{IMPORTS_MTD},$headNumber);
	$worksheet->write($a,8, $s->{TOTAL_MTD},$headNumber);
	$worksheet->write($a,9, $s->{CONT_MTD},$headPct);
	
	$worksheet->write($a,10, $s->{IMPORTS_LY_MTD},$headNumber);					
	$worksheet->write($a,11, $s->{TOTAL_LY_MTD},$headNumber);
	$worksheet->write($a,12, $s->{CONT_LY_MTD},$headPct);
	
	$worksheet->write($a,13, "",$headPct);
	
	$worksheet->write($a,14, $s->{IMPORTS_QTD},$headNumber);						
	$worksheet->write($a,15, $s->{TOTAL_QTD},$headNumber);
	$worksheet->write($a,16, $s->{CONT_QTD},$headPct);
	
	$worksheet->write($a,17, $s->{IMPORTS_LY_QTD},$headNumber);						
	$worksheet->write($a,18, $s->{TOTAL_LY_QTD},$headNumber);
	$worksheet->write($a,19, $s->{CONT_LY_QTD},$headPct);
	
	$worksheet->write($a,20, "",$headPct);
	
	$a++; #INCREMENT VARIABLE a
}
	
	$worksheet->write($a,7, $total_imports_mtd, $headNumber);
	$worksheet->write($a,8, $total_total_mtd, $headNumber);
		if ($total_total_mtd eq 0) { $worksheet->write($a,9, "", $headPct); }
		else { $worksheet->write($a,9, $total_imports_mtd/$total_total_mtd, $headPct); }
	
	$worksheet->write($a,10, $total_imports_ly_mtd, $headNumber);					
	$worksheet->write($a,11, $total_total_ly_mtd, $headNumber);
		if ($total_total_ly_mtd eq 0) { $worksheet->write($a,12, "", $headPct); }
		else { $worksheet->write($a,12, $total_imports_ly_mtd/$total_total_ly_mtd, $headPct); }
		
	$worksheet->write($a,13, "",$headPct);
	
	$worksheet->write($a,14, $total_imports_qtd, $headNumber);						
	$worksheet->write($a,15, $total_total_qtd, $headNumber);
		if ($total_total_qtd eq 0) { $worksheet->write($a,16, "", $headPct); }
		else { $worksheet->write($a,16, $total_imports_qtd/$total_total_qtd, $headPct); }
	
	$worksheet->write($a,17, $total_imports_ly_qtd, $headNumber);						
	$worksheet->write($a,18, $total_total_ly_qtd, $headNumber);
		if ($total_total_ly_qtd eq 0) { $worksheet->write($a,19, "", $headPct); }
		else { $worksheet->write($a,19, $total_imports_ly_qtd/$total_total_ly_qtd, $headPct); }
	
	$worksheet->write($a,20, "",$headPct);
	
$worksheet->write($loc, 2, $loc_code . " - " . $loc_desc, $bold);			
$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}


sub query_dept_country {
	
$sls1 = $dbh->prepare (qq{
			SELECT merch_group_code, merch_group_desc, 
			  SUM(D_IMPORTS_MTD) AS D_IMPORTS_MTD, SUM(D_IMPORTS_LY_MTD) AS D_IMPORTS_LY_MTD, 
			  CASE WHEN SUM(D_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_MTD) / SUM(D_IMPORTS_LY_MTD) ),1) END AS D_MTD,
			  SUM(D_IMPORTS_QTD) AS D_IMPORTS_QTD, SUM(D_IMPORTS_LY_QTD) D_IMPORTS_LY_QTD,
			  CASE WHEN SUM(D_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_QTD) / SUM(D_IMPORTS_LY_QTD) ),1) END AS Q_MTD,  
			  SUM(L_IMPORTS_MTD) AS L_IMPORTS_MTD, SUM(L_IMPORTS_LY_MTD) AS L_IMPORTS_LY_MTD,   
			  CASE WHEN SUM(L_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(L_IMPORTS_MTD) / SUM(L_IMPORTS_LY_MTD) ),1) END AS L_MTD,
			  SUM(L_IMPORTS_QTD) AS L_IMPORTS_QTD, SUM(L_IMPORTS_LY_QTD) L_IMPORTS_LY_QTD,
			  CASE WHEN SUM(L_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(L_IMPORTS_QTD) / SUM(L_IMPORTS_LY_QTD) ),1) END AS Q_MTD, 
			  SUM(D_IMPORTS_MTD+L_IMPORTS_MTD) AS IMPORTS_MTD, SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) AS IMPORTS_LY_MTD,  
			  CASE WHEN SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_MTD+L_IMPORTS_MTD) / SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) ),1) END AS MTD,
			  SUM(D_IMPORTS_QTD+L_IMPORTS_QTD) AS IMPORTS_QTD, SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) AS IMPORTS_LY_QTD,
			  CASE WHEN SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_QTD+L_IMPORTS_QTD) / SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) ),1) END AS QTD
			FROM
				(SELECT merch_group_code, merch_group_desc, SUM(TOTAL_RETAIL_MTD)/1000 D_IMPORTS_MTD, SUM(TOTAL_RETAIL_LY_MTD)/1000 D_IMPORTS_LY_MTD,
											SUM(TOTAL_RETAIL_QTD)/1000 D_IMPORTS_QTD, SUM(TOTAL_RETAIL_LY_QTD)/1000 D_IMPORTS_LY_QTD,
											0 AS L_IMPORTS_MTD, 0 AS L_IMPORTS_LY_MTD,
											0 AS L_IMPORTS_QTD, 0 AS L_IMPORTS_LY_QTD
					FROM METRO_IT_VENDOR_IMPORTS4
					WHERE COUNTRY IS NOT NULL AND IMP_TYPE = 'Direct'
					GROUP BY merch_group_code, merch_group_desc
					UNION ALL
				SELECT merch_group_code, merch_group_desc,  0 AS D_IMPORTS_MTD, 0 AS D_IMPORTS_LY_MTD,
												0 AS D_IMPORTS_QTD, 0 AS D_IMPORTS_LY_QTD,
											SUM(TOTAL_RETAIL_MTD)/1000 L_IMPORTS_MTD, SUM(TOTAL_RETAIL_LY_MTD)/1000 L_IMPORTS_LY_MTD,
											SUM(TOTAL_RETAIL_QTD)/1000 L_IMPORTS_QTD, SUM(TOTAL_RETAIL_LY_QTD)/1000 L_IMPORTS_LY_QTD
					FROM METRO_IT_VENDOR_IMPORTS4
					WHERE COUNTRY IS NOT NULL AND IMP_TYPE = 'Local'
					GROUP BY merch_group_code, merch_group_desc)
			GROUP BY merch_group_code, merch_group_desc ORDER BY 1
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{MERCH_GROUP_CODE};
	$merch_group_desc = $s->{MERCH_GROUP_DESC};
	
	$sls2 = $dbh->prepare (qq{
			SELECT group_code, group_desc, 
			  SUM(D_IMPORTS_MTD) AS D_IMPORTS_MTD, SUM(D_IMPORTS_LY_MTD) AS D_IMPORTS_LY_MTD, 
			  CASE WHEN SUM(D_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_MTD) / SUM(D_IMPORTS_LY_MTD) ),1) END AS D_MTD,
			  SUM(D_IMPORTS_QTD) AS D_IMPORTS_QTD, SUM(D_IMPORTS_LY_QTD) D_IMPORTS_LY_QTD,
			  CASE WHEN SUM(D_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_QTD) / SUM(D_IMPORTS_LY_QTD) ),1) END AS Q_MTD,  
			  SUM(L_IMPORTS_MTD) AS L_IMPORTS_MTD, SUM(L_IMPORTS_LY_MTD) AS L_IMPORTS_LY_MTD,   
			  CASE WHEN SUM(L_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(L_IMPORTS_MTD) / SUM(L_IMPORTS_LY_MTD) ),1) END AS L_MTD,
			  SUM(L_IMPORTS_QTD) AS L_IMPORTS_QTD, SUM(L_IMPORTS_LY_QTD) L_IMPORTS_LY_QTD,
			  CASE WHEN SUM(L_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(L_IMPORTS_QTD) / SUM(L_IMPORTS_LY_QTD) ),1) END AS Q_MTD, 
			  SUM(D_IMPORTS_MTD+L_IMPORTS_MTD) AS IMPORTS_MTD, SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) AS IMPORTS_LY_MTD,  
			  CASE WHEN SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_MTD+L_IMPORTS_MTD) / SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) ),1) END AS MTD,
			  SUM(D_IMPORTS_QTD+L_IMPORTS_QTD) AS IMPORTS_QTD, SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) AS IMPORTS_LY_QTD,
			  CASE WHEN SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_QTD+L_IMPORTS_QTD) / SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) ),1) END AS QTD
			FROM
				(SELECT group_code, group_desc, SUM(TOTAL_RETAIL_MTD)/1000 D_IMPORTS_MTD, SUM(TOTAL_RETAIL_LY_MTD)/1000 D_IMPORTS_LY_MTD,
											SUM(TOTAL_RETAIL_QTD)/1000 D_IMPORTS_QTD, SUM(TOTAL_RETAIL_LY_QTD)/1000 D_IMPORTS_LY_QTD,
											0 AS L_IMPORTS_MTD, 0 AS L_IMPORTS_LY_MTD,
											0 AS L_IMPORTS_QTD, 0 AS L_IMPORTS_LY_QTD
					FROM METRO_IT_VENDOR_IMPORTS4
					WHERE merch_group_code = '$merch_group_code' and COUNTRY IS NOT NULL AND IMP_TYPE = 'Direct'
					GROUP BY group_code, group_desc
					UNION ALL
				SELECT group_code, group_desc,  0 AS D_IMPORTS_MTD, 0 AS D_IMPORTS_LY_MTD,
												0 AS D_IMPORTS_QTD, 0 AS D_IMPORTS_LY_QTD,
											SUM(TOTAL_RETAIL_MTD)/1000 L_IMPORTS_MTD, SUM(TOTAL_RETAIL_LY_MTD)/1000 L_IMPORTS_LY_MTD,
											SUM(TOTAL_RETAIL_QTD)/1000 L_IMPORTS_QTD, SUM(TOTAL_RETAIL_LY_QTD)/1000 L_IMPORTS_LY_QTD
					FROM METRO_IT_VENDOR_IMPORTS4
					WHERE merch_group_code = '$merch_group_code' and COUNTRY IS NOT NULL AND IMP_TYPE = 'Local'
					GROUP BY group_code, group_desc)
			GROUP BY group_code, group_desc ORDER BY 1
							 });	
	$sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$group_code = $s->{GROUP_CODE};
		$group_desc = $s->{GROUP_DESC};
				
		$sls3 = $dbh->prepare (qq{
			SELECT division, division_desc, 
			  SUM(D_IMPORTS_MTD) AS D_IMPORTS_MTD, SUM(D_IMPORTS_LY_MTD) AS D_IMPORTS_LY_MTD, 
			  CASE WHEN SUM(D_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_MTD) / SUM(D_IMPORTS_LY_MTD) ),1) END AS D_MTD,
			  SUM(D_IMPORTS_QTD) AS D_IMPORTS_QTD, SUM(D_IMPORTS_LY_QTD) D_IMPORTS_LY_QTD,
			  CASE WHEN SUM(D_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_QTD) / SUM(D_IMPORTS_LY_QTD) ),1) END AS Q_MTD,  
			  SUM(L_IMPORTS_MTD) AS L_IMPORTS_MTD, SUM(L_IMPORTS_LY_MTD) AS L_IMPORTS_LY_MTD,   
			  CASE WHEN SUM(L_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(L_IMPORTS_MTD) / SUM(L_IMPORTS_LY_MTD) ),1) END AS L_MTD,
			  SUM(L_IMPORTS_QTD) AS L_IMPORTS_QTD, SUM(L_IMPORTS_LY_QTD) L_IMPORTS_LY_QTD,
			  CASE WHEN SUM(L_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(L_IMPORTS_QTD) / SUM(L_IMPORTS_LY_QTD) ),1) END AS Q_MTD, 
			  SUM(D_IMPORTS_MTD+L_IMPORTS_MTD) AS IMPORTS_MTD, SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) AS IMPORTS_LY_MTD,  
			  CASE WHEN SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_MTD+L_IMPORTS_MTD) / SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) ),1) END AS MTD,
			  SUM(D_IMPORTS_QTD+L_IMPORTS_QTD) AS IMPORTS_QTD, SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) AS IMPORTS_LY_QTD,
			  CASE WHEN SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_QTD+L_IMPORTS_QTD) / SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) ),1) END AS QTD
			FROM
				(SELECT division, division_desc, SUM(TOTAL_RETAIL_MTD)/1000 D_IMPORTS_MTD, SUM(TOTAL_RETAIL_LY_MTD)/1000 D_IMPORTS_LY_MTD,
											SUM(TOTAL_RETAIL_QTD)/1000 D_IMPORTS_QTD, SUM(TOTAL_RETAIL_LY_QTD)/1000 D_IMPORTS_LY_QTD,
											0 AS L_IMPORTS_MTD, 0 AS L_IMPORTS_LY_MTD,
											0 AS L_IMPORTS_QTD, 0 AS L_IMPORTS_LY_QTD
					FROM METRO_IT_VENDOR_IMPORTS4
					WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and COUNTRY IS NOT NULL AND IMP_TYPE = 'Direct'
					GROUP BY division, division_desc
					UNION ALL
				SELECT division, division_desc,  0 AS D_IMPORTS_MTD, 0 AS D_IMPORTS_LY_MTD,
												0 AS D_IMPORTS_QTD, 0 AS D_IMPORTS_LY_QTD,
											SUM(TOTAL_RETAIL_MTD)/1000 L_IMPORTS_MTD, SUM(TOTAL_RETAIL_LY_MTD)/1000 L_IMPORTS_LY_MTD,
											SUM(TOTAL_RETAIL_QTD)/1000 L_IMPORTS_QTD, SUM(TOTAL_RETAIL_LY_QTD)/1000 L_IMPORTS_LY_QTD
					FROM METRO_IT_VENDOR_IMPORTS4
					WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and COUNTRY IS NOT NULL AND IMP_TYPE = 'Local'
					GROUP BY division, division_desc)
			GROUP BY division, division_desc ORDER BY 1
								});
		$sls3->execute();
		
		$grp_counter = $a;
		while(my $s = $sls3->fetchrow_hashref()){
			$division = $s->{DIVISION};
			$division_desc = $s->{DIVISION_DESC};
			
			$sls4 = $dbh->prepare (qq{					
			SELECT country, 
			  SUM(D_IMPORTS_MTD) AS D_IMPORTS_MTD, SUM(D_IMPORTS_LY_MTD) AS D_IMPORTS_LY_MTD, 
			  CASE WHEN SUM(D_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_MTD) / SUM(D_IMPORTS_LY_MTD) ),1) END AS D_MTD,
			  SUM(D_IMPORTS_QTD) AS D_IMPORTS_QTD, SUM(D_IMPORTS_LY_QTD) D_IMPORTS_LY_QTD,
			  CASE WHEN SUM(D_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_QTD) / SUM(D_IMPORTS_LY_QTD) ),1) END AS Q_MTD,  
			  SUM(L_IMPORTS_MTD) AS L_IMPORTS_MTD, SUM(L_IMPORTS_LY_MTD) AS L_IMPORTS_LY_MTD,   
			  CASE WHEN SUM(L_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(L_IMPORTS_MTD) / SUM(L_IMPORTS_LY_MTD) ),1) END AS L_MTD,
			  SUM(L_IMPORTS_QTD) AS L_IMPORTS_QTD, SUM(L_IMPORTS_LY_QTD) L_IMPORTS_LY_QTD,
			  CASE WHEN SUM(L_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(L_IMPORTS_QTD) / SUM(L_IMPORTS_LY_QTD) ),1) END AS Q_MTD, 
			  SUM(D_IMPORTS_MTD+L_IMPORTS_MTD) AS IMPORTS_MTD, SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) AS IMPORTS_LY_MTD,  
			  CASE WHEN SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_MTD+L_IMPORTS_MTD) / SUM(D_IMPORTS_LY_MTD+L_IMPORTS_LY_MTD) ),1) END AS MTD,
			  SUM(D_IMPORTS_QTD+L_IMPORTS_QTD) AS IMPORTS_QTD, SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) AS IMPORTS_LY_QTD,
			  CASE WHEN SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) = 0 THEN NULL ELSE ROUND(( SUM(D_IMPORTS_QTD+L_IMPORTS_QTD) / SUM(D_IMPORTS_LY_QTD+L_IMPORTS_LY_QTD) ),1) END AS QTD
			FROM
				(SELECT country, SUM(TOTAL_RETAIL_MTD)/1000 D_IMPORTS_MTD, SUM(TOTAL_RETAIL_LY_MTD)/1000 D_IMPORTS_LY_MTD,
											SUM(TOTAL_RETAIL_QTD)/1000 D_IMPORTS_QTD, SUM(TOTAL_RETAIL_LY_QTD)/1000 D_IMPORTS_LY_QTD,
											0 AS L_IMPORTS_MTD, 0 AS L_IMPORTS_LY_MTD,
											0 AS L_IMPORTS_QTD, 0 AS L_IMPORTS_LY_QTD
					FROM METRO_IT_VENDOR_IMPORTS4
					WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and division = '$division' and COUNTRY IS NOT NULL AND IMP_TYPE = 'Direct'
					GROUP BY country
					UNION ALL
				SELECT country,  0 AS D_IMPORTS_MTD, 0 AS D_IMPORTS_LY_MTD,
												0 AS D_IMPORTS_QTD, 0 AS D_IMPORTS_LY_QTD,
											SUM(TOTAL_RETAIL_MTD)/1000 L_IMPORTS_MTD, SUM(TOTAL_RETAIL_LY_MTD)/1000 L_IMPORTS_LY_MTD,
											SUM(TOTAL_RETAIL_QTD)/1000 L_IMPORTS_QTD, SUM(TOTAL_RETAIL_LY_QTD)/1000 L_IMPORTS_LY_QTD
					FROM METRO_IT_VENDOR_IMPORTS4
					WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and division = '$division' and COUNTRY IS NOT NULL AND IMP_TYPE = 'Local'
					GROUP BY country)
			GROUP BY country ORDER BY 1
									});
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){
				
				$worksheet->merge_range($a, 5, $a, 6, $s->{COUNTRY},$desc);
				
				$worksheet->write($a,7, $s->{D_IMPORTS_MTD},$border1);
				$worksheet->write($a,8, $s->{D_IMPORTS_LY_MTD},$border1);
				$worksheet->write($a,9, $s->{D_MTD} . " %",$border1);
				
				$worksheet->write($a,10, $s->{L_IMPORTS_MTD},$border1);					
				$worksheet->write($a,11, $s->{L_IMPORTS_LY_MTD},$border1);
				$worksheet->write($a,12, $s->{L_MTD} . " %",$border1);
				
				$worksheet->write($a,13, $s->{IMPORTS_MTD},$border1);						
				$worksheet->write($a,14, $s->{IMPORTS_LY_MTD},$border1);
				$worksheet->write($a,15, $s->{MTD} . " %",$border1);
				
				$worksheet->write($a,16, $s->{D_IMPORTS_QTD},$border1);						
				$worksheet->write($a,17, $s->{D_IMPORTS_LY_QTD},$border1);
				$worksheet->write($a,18, $s->{Q_MTD} . " %",$border1);
				
				$worksheet->write($a,19, $s->{L_IMPORTS_QTD},$border1);						
				$worksheet->write($a,20, $s->{L_IMPORTS_LY_QTD},$border1);
				$worksheet->write($a,21, $s->{Q_MTD} . " %",$border1);
				
				$worksheet->write($a,22, $s->{IMPORTS_QTD},$border1);						
				$worksheet->write($a,23, $s->{IMPORTS_LY_QTD},$border1);
				$worksheet->write($a,24, $s->{QTD} . " %",$border1);
				
				$a++;
				$counter++;
		
			}			
			
			$worksheet->write($a,7, $s->{D_IMPORTS_MTD},$bodyNum);
			$worksheet->write($a,8, $s->{D_IMPORTS_LY_MTD},$bodyNum);
			$worksheet->write($a,9, $s->{D_MTD} . " %",$bodyNum);
			
			$worksheet->write($a,10, $s->{L_IMPORTS_MTD},$bodyNum);					
			$worksheet->write($a,11, $s->{L_IMPORTS_LY_MTD},$bodyNum);
			$worksheet->write($a,12, $s->{L_MTD} . " %",$bodyNum);
			
			$worksheet->write($a,13, $s->{IMPORTS_MTD},$bodyNum);						
			$worksheet->write($a,14, $s->{IMPORTS_LY_MTD},$bodyNum);
			$worksheet->write($a,15, $s->{MTD} . " %",$bodyNum);
			
			$worksheet->write($a,16, $s->{D_IMPORTS_QTD},$bodyNum);						
			$worksheet->write($a,17, $s->{D_IMPORTS_LY_QTD},$bodyNum);
			$worksheet->write($a,18, $s->{Q_MTD} . " %",$bodyNum);
			
			$worksheet->write($a,19, $s->{L_IMPORTS_QTD},$bodyNum);						
			$worksheet->write($a,20, $s->{L_IMPORTS_LY_QTD},$bodyNum);
			$worksheet->write($a,21, $s->{Q_MTD} . " %",$bodyNum);
			
			$worksheet->write($a,22, $s->{IMPORTS_QTD},$bodyNum);						
			$worksheet->write($a,23, $s->{IMPORTS_LY_QTD},$bodyNum);
			$worksheet->write($a,24, $s->{QTD} . " %",$bodyNum);

			$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
			
			$counter = 0; #RESET dept_counter
			$a++; #INCREMENT VARIABLE a
		}
		
		$worksheet->write($a,7, $s->{D_IMPORTS_MTD},$bodyNum);
		$worksheet->write($a,8, $s->{D_IMPORTS_LY_MTD},$bodyNum);
		$worksheet->write($a,9, $s->{D_MTD} . " %",$bodyNum);
		
		$worksheet->write($a,10, $s->{L_IMPORTS_MTD},$bodyNum);					
		$worksheet->write($a,11, $s->{L_IMPORTS_LY_MTD},$bodyNum);
		$worksheet->write($a,12, $s->{L_MTD} . " %",$bodyNum);
		
		$worksheet->write($a,13, $s->{IMPORTS_MTD},$bodyNum);						
		$worksheet->write($a,14, $s->{IMPORTS_LY_MTD},$bodyNum);
		$worksheet->write($a,15, $s->{MTD} . " %",$bodyNum);
		
		$worksheet->write($a,16, $s->{D_IMPORTS_QTD},$bodyNum);						
		$worksheet->write($a,17, $s->{D_IMPORTS_LY_QTD},$bodyNum);
		$worksheet->write($a,18, $s->{Q_MTD} . " %",$bodyNum);
		
		$worksheet->write($a,19, $s->{L_IMPORTS_QTD},$bodyNum);						
		$worksheet->write($a,20, $s->{L_IMPORTS_LY_QTD},$bodyNum);
		$worksheet->write($a,21, $s->{Q_MTD} . " %",$bodyNum);
		
		$worksheet->write($a,22, $s->{IMPORTS_QTD},$bodyNum);						
		$worksheet->write($a,23, $s->{IMPORTS_LY_QTD},$bodyNum);
		$worksheet->write($a,24, $s->{QTD} . " %",$bodyNum);
		
		$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

		$a++; #INCREMENT VARIABLE a
	}
	
	$total_d_imports_mtd += $s->{D_IMPORTS_MTD};
	$total_d_imports_ly_mtd += $s->{D_IMPORTS_LY_MTD};
	
	$total_l_imports_mtd += $s->{L_IMPORTS_MTD};
	$total_l_imports_ly_mtd += $s->{L_IMPORTS_LY_MTD};
			
	$total_imports_mtd += $s->{IMPORTS_MTD};
	$total_imports_ly_mtd += $s->{IMPORTS_LY_MTD};
			
	$total_d_imports_qtd += $s->{D_IMPORTS_QTD};
	$total_d_imports_ly_qtd += $s->{D_IMPORTS_LY_QTD};
	
	$total_l_imports_qtd += $s->{L_IMPORTS_QTD};
	$total_l_imports_ly_qtd += $s->{L_IMPORTS_LY_QTD};
			
	$total_imports_qtd += $s->{IMPORTS_QTD};
	$total_imports_ly_qtd += $s->{IMPORTS_LY_QTD};
	
	if ($merch_group_code eq 'DS'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
	}
	
	elsif($merch_group_code eq 'SU'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
	}
	
	elsif($merch_group_code eq 'Z_OT'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Others', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'OTHERS', $border2 );
	}
		
	$worksheet->write($a,7, $s->{D_IMPORTS_MTD},$headNumber);
	$worksheet->write($a,8, $s->{D_IMPORTS_LY_MTD},$headNumber);
	$worksheet->write($a,9, $s->{D_MTD} . " %",$headNumber);
			
	$worksheet->write($a,10, $s->{L_IMPORTS_MTD},$headNumber);					
	$worksheet->write($a,11, $s->{L_IMPORTS_LY_MTD},$headNumber);
	$worksheet->write($a,12, $s->{L_MTD} . " %",$headNumber);
			
	$worksheet->write($a,13, $s->{IMPORTS_MTD},$headNumber);						
	$worksheet->write($a,14, $s->{IMPORTS_LY_MTD},$headNumber);
	$worksheet->write($a,15, $s->{MTD} . " %",$headNumber);
			
	$worksheet->write($a,16, $s->{D_IMPORTS_QTD},$headNumber);						
	$worksheet->write($a,17, $s->{D_IMPORTS_LY_QTD},$headNumber);
	$worksheet->write($a,18, $s->{Q_MTD} . " %",$headNumber);
			
	$worksheet->write($a,19, $s->{L_IMPORTS_QTD},$headNumber);						
	$worksheet->write($a,20, $s->{L_IMPORTS_LY_QTD},$headNumber);
	$worksheet->write($a,21, $s->{Q_MTD} . " %",$headNumber);
		
	$worksheet->write($a,22, $s->{IMPORTS_QTD},$headNumber);						
	$worksheet->write($a,23, $s->{IMPORTS_LY_QTD},$headNumber);
	$worksheet->write($a,24, $s->{QTD} . " %",$headNumber);
	
	$a++; #INCREMENT VARIABLE a
}
			
	$worksheet->write($a,7, $total_d_imports_mtd, $headNumber);
	$worksheet->write($a,8, $total_d_imports_ly_mtd, $headNumber);
		if ($total_d_imports_ly_mtd <= 0){
			$worksheet->write($a,9, "", $headPct); }
		else{
			$worksheet->write($a,9, ($total_d_imports_mtd)/$total_d_imports_ly_mtd   . " %", $headNumber); }
			
	$worksheet->write($a,10, $total_l_imports_mtd, $headNumber);					
	$worksheet->write($a,11, $total_l_imports_ly_mtd, $headNumber);
		if ($total_l_imports_ly_mtd <= 0){
			$worksheet->write($a,12, "", $headPct); }
		else{
			$worksheet->write($a,12, ($total_l_imports_mtd)/$total_l_imports_ly_mtd  . " %", $headNumber); }
			
	$worksheet->write($a,13, $total_imports_mtd, $headNumber);						
	$worksheet->write($a,14, $total_imports_ly_mtd, $headNumber);
		if ($total_imports_ly_mtd <= 0){
			$worksheet->write($a,15, "", $headPct); }
		else{
			$worksheet->write($a,15, ($total_imports_mtd)/$total_imports_ly_mtd  . " %", $headNumber); }
			
	$worksheet->write($a,16, $total_d_imports_qtd, $headNumber);						
	$worksheet->write($a,17, $total_d_imports_ly_qtd, $headNumber);
		if ($total_d_imports_ly_qtd <= 0){
			$worksheet->write($a,18, "", $headPct); }
		else{
			$worksheet->write($a,18, ($total_d_imports_qtd)/$total_d_imports_ly_qtd  . " %", $headNumber); }
			
	$worksheet->write($a,19, $total_l_imports_qtd, $headNumber);						
	$worksheet->write($a,20, $total_l_imports_ly_qtd, $headNumber);
		if ($total_l_imports_ly_qtd <= 0){
			$worksheet->write($a,21, "",$headPct); }
		else{
			$worksheet->write($a,21, ($total_l_imports_qtd)/$total_l_imports_ly_qtd  . " %", $headNumber); }
		
	$worksheet->write($a,22, $total_imports_qtd,$headNumber);						
	$worksheet->write($a,23, $total_imports_ly_qtd,$headNumber);
		if ($total_imports_ly_qtd <= 0){
			$worksheet->write($a,24, "",$headPct); }
		else{
			$worksheet->write($a,24, ($total_imports_qtd)/$total_imports_ly_qtd  . " %", $headNumber); }
	
$worksheet->write($loc, 2, $loc_desc, $bold);
$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}


sub query_dept_vendor {
	
$sls1 = $dbh->prepare (qq{SELECT merch_group_code, merch_group_desc, 
							  SUM(TOTAL_RETAIL_MTD)/1000 TOTAL_MTD, 
							  SUM(TOTAL_RETAIL_LY_MTD)/1000 TOTAL_LY_MTD, 
							  CASE WHEN SUM(TOTAL_RETAIL_LY_MTD) = 0 THEN NULL ELSE ROUND(( (SUM(TOTAL_RETAIL_MTD)/1000) / (SUM(TOTAL_RETAIL_LY_MTD)/1000) ),1) END AS CONT_MTD,
							  SUM(TOTAL_RETAIL_QTD)/1000 TOTAL_QTD, 
							  SUM(TOTAL_RETAIL_LY_QTD)/1000 TOTAL_LY_QTD,
							  CASE WHEN SUM(TOTAL_RETAIL_LY_QTD) = 0 THEN NULL ELSE ROUND(( (SUM(TOTAL_RETAIL_QTD)/1000) / (SUM(TOTAL_RETAIL_LY_QTD)/1000) ),1) END AS CONT_QTD
						FROM metro_it_vendor_imports4
						GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC
						ORDER BY 1
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{MERCH_GROUP_CODE};
	$merch_group_desc = $s->{MERCH_GROUP_DESC};
	
	$sls2 = $dbh->prepare (qq{SELECT group_code, group_desc, 
								  SUM(TOTAL_RETAIL_MTD)/1000 TOTAL_MTD, 
								  SUM(TOTAL_RETAIL_LY_MTD)/1000 TOTAL_LY_MTD, 
								  CASE WHEN SUM(TOTAL_RETAIL_LY_MTD) = 0 THEN NULL ELSE ROUND(( (SUM(TOTAL_RETAIL_MTD)/1000) / (SUM(TOTAL_RETAIL_LY_MTD)/1000) ),1) END AS CONT_MTD,
								  SUM(TOTAL_RETAIL_QTD)/1000 TOTAL_QTD, 
								  SUM(TOTAL_RETAIL_LY_QTD)/1000 TOTAL_LY_QTD,
								  CASE WHEN SUM(TOTAL_RETAIL_LY_QTD) = 0 THEN NULL ELSE ROUND(( (SUM(TOTAL_RETAIL_QTD)/1000) / (SUM(TOTAL_RETAIL_LY_QTD)/1000) ),1) END AS CONT_QTD
							FROM metro_it_vendor_imports4
							WHERE merch_group_code = '$merch_group_code'
							GROUP BY group_code, group_desc
							ORDER BY 1
							 });	
	$sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$group_code = $s->{GROUP_CODE};
		$group_desc = $s->{GROUP_DESC};
				
		$sls3 = $dbh->prepare (qq{SELECT division, division_desc, 
									  SUM(TOTAL_RETAIL_MTD)/1000 TOTAL_MTD, 
									  SUM(TOTAL_RETAIL_LY_MTD)/1000 TOTAL_LY_MTD, 
									 CASE WHEN SUM(TOTAL_RETAIL_LY_MTD)=0 THEN NULL ELSE ROUND(( (SUM(TOTAL_RETAIL_MTD)/1000) / (SUM(TOTAL_RETAIL_LY_MTD)/1000) ),1) END AS CONT_MTD,
									  SUM(TOTAL_RETAIL_QTD)/1000 TOTAL_QTD, 
									  SUM(TOTAL_RETAIL_LY_QTD)/1000 TOTAL_LY_QTD,
									 CASE WHEN SUM(TOTAL_RETAIL_LY_QTD)=0 THEN NULL ELSE ROUND(( (SUM(TOTAL_RETAIL_QTD)/1000) / (SUM(TOTAL_RETAIL_LY_QTD)/1000) ),1) END AS CONT_QTD
								FROM metro_it_vendor_imports4
								WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code'
								GROUP BY division, division_desc
								ORDER BY 1
								});
		$sls3->execute();
		
		$grp_counter = $a;
		while(my $s = $sls3->fetchrow_hashref()){
			$division = $s->{DIVISION};
			$division_desc = $s->{DIVISION_DESC};
			
			$sls4 = $dbh->prepare (qq{SELECT supplier, sup_name, 
										  SUM(TOTAL_RETAIL_MTD)/1000 TOTAL_MTD, 
										  SUM(TOTAL_RETAIL_LY_MTD)/1000 TOTAL_LY_MTD, 
									CASE WHEN SUM(TOTAL_RETAIL_LY_MTD)=0 THEN NULL ELSE ROUND(( (SUM(TOTAL_RETAIL_MTD)/1000) / (SUM(TOTAL_RETAIL_LY_MTD)/1000) ),1) END AS CONT_MTD,
										  SUM(TOTAL_RETAIL_QTD)/1000 TOTAL_QTD, 
										  SUM(TOTAL_RETAIL_LY_QTD)/1000 TOTAL_LY_QTD,
									CASE WHEN SUM(TOTAL_RETAIL_LY_QTD)=0 THEN NULL ELSE ROUND(( (SUM(TOTAL_RETAIL_QTD)/1000) / (SUM(TOTAL_RETAIL_LY_QTD)/1000) ),1) END AS CONT_QTD
									FROM metro_it_vendor_imports4
									WHERE group_code = '$group_code' and merch_group_code = '$merch_group_code' and division = '$division'
									GROUP BY supplier, sup_name
									ORDER BY 1	
									});
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){
				
				$worksheet->write($a,5, $s->{SUPPLIER},$desc);
				$worksheet->write($a,6, $s->{SUP_NAME},$desc);
				
				$worksheet->write($a,7, $s->{TOTAL_MTD},$border1);				
				$worksheet->write($a,8, $s->{TOTAL_LY_MTD},$border1);
				
				$worksheet->write($a,9, $s->{CONT_MTD} . " %",$border1);
				
				$worksheet->write($a,10, $s->{TOTAL_QTD},$border1);
				$worksheet->write($a,11, $s->{TOTAL_LY_QTD},$border1);
				
				$worksheet->write($a,12, $s->{CONT_QTD} . " %",$border1);
				
				$a++;
				$counter++;
		
			}			
			
			$worksheet->write($a,7, $s->{TOTAL_MTD},$bodyNum);				
			$worksheet->write($a,8, $s->{TOTAL_LY_MTD},$bodyNum);
			
			$worksheet->write($a,9, $s->{CONT_MTD} . " %",$bodyNum);
			
			$worksheet->write($a,10, $s->{TOTAL_QTD},$bodyNum);
			$worksheet->write($a,11, $s->{TOTAL_LY_QTD},$bodyNum);
			
			$worksheet->write($a,12, $s->{CONT_QTD} . " %",$bodyNum);

			$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
			
			$counter = 0; #RESET dept_counter
			$a++; #INCREMENT VARIABLE a
		}
		
		$worksheet->write($a,7, $s->{TOTAL_MTD},$bodyNum);				
		$worksheet->write($a,8, $s->{TOTAL_LY_MTD},$bodyNum);
		
		$worksheet->write($a,9, $s->{CONT_MTD} . " %",$bodyNum);
		
		$worksheet->write($a,10, $s->{TOTAL_QTD},$bodyNum);
		$worksheet->write($a,11, $s->{TOTAL_LY_QTD},$bodyNum);
		
		$worksheet->write($a,12, $s->{CONT_QTD} . " %",$bodyNum);
		
		$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

		$a++; #INCREMENT VARIABLE a
	}
	$total_total_mtd += $s->{TOTAL_MTD};
	$total_total_ly_mtd += $s->{TOTAL_LY_MTD};
			
	$total_total_qtd += $s->{TOTAL_QTD};
	$total_total_ly_qtd += $s->{TOTAL_LY_QTD};
	
	if ($merch_group_code eq 'DS'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
	}
	
	elsif($merch_group_code eq 'SU'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
	}
	
	elsif($merch_group_code eq 'Z_OT'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Others', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'OTHERS', $border2 );
	}
	
	$worksheet->write($a,7, $s->{TOTAL_MTD},$headNumber);				
	$worksheet->write($a,8, $s->{TOTAL_LY_MTD},$headNumber);
	
	$worksheet->write($a,9, $s->{CONT_MTD} . " %",$headNumber);
			
	$worksheet->write($a,10, $s->{TOTAL_QTD},$headNumber);
	$worksheet->write($a,11, $s->{TOTAL_LY_QTD},$headNumber);
	
	$worksheet->write($a,12, $s->{CONT_QTD} . " %",$headNumber);
	
	$a++; #INCREMENT VARIABLE a
}
	
	$worksheet->write($a,7, $total_total_mtd, $headNumber);					
	$worksheet->write($a,8, $total_total_ly_mtd, $headNumber);
		if ($total_total_ly_mtd eq 0) { $worksheet->write($a,9, "", $headPct); }
		else { $worksheet->write($a,9, $total_total_mtd/$total_total_ly_mtd  . " %", $headNumber); }
		
	$worksheet->write($a,10, $total_total_qtd, $headNumber);
	$worksheet->write($a,11, $total_total_ly_qtd, $headNumber);
		if ($total_total_ly_qtd eq 0) { $worksheet->write($a,12, "", $headPct); }
		else { $worksheet->write($a,12, $total_total_qtd/$total_total_ly_qtd  . " %", $headNumber); }
		
$worksheet->write($loc, 2, $loc_desc, $bold);
$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}


sub calc8 { 

if($type_test eq 3 or $type_test eq 4){
	foreach my $col( 7, 8, 10, 12, 13, 15, 17, 18, 20, 22, 23, 25, 27, 28, 30, 32, 33, 35 ){
		if($s1f2_row ne 0){
			my $sum = '=(SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . '))-'. xl_rowcol_to_cell( $s1f2_row, $col );
				$worksheet->write( $a, $col, $sum, $bodyNum );
			
			if($col eq 8 or $col eq 13 or $col eq 18 or $col eq 23 or $col eq 28 or $col eq 33){
				if ($wtd_net_ly eq 0 or $mtd_net_ly eq 0 or $qtd_net_ly eq 0) {
						$worksheet->write( $a, $col+1, "", $bodyPct );}
				else{
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );}	
			}
			elsif ($col eq 10 or $col eq 15 or $col eq 20 or $col eq 25 or $col eq 30 or $col eq 35){
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
			}
				
		}
		else{
			my $sum = '=SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . ')';
				$worksheet->write( $a, $col, $sum, $bodyNum );

			if($col eq 8 or $col eq 13 or $col eq 18 or $col eq 23 or $col eq 28 or $col eq 33){
				if ($wtd_net_ly eq 0 or $mtd_net_ly eq 0 or $qtd_net_ly eq 0) {
						$worksheet->write( $a, $col+1, "", $bodyPct );}
				else{
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );}	
			}
			elsif ($col eq 10 or $col eq 15 or $col eq 20 or $col eq 25 or $col eq 30 or $col eq 35){
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
			}
		}		
	}
}

else{
	foreach my $col( 7, 8, 10, 12, 13, 15, 17, 18, 20 ){
		if($s1f2_row ne 0){
			my $sum = '=(SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . '))-'. xl_rowcol_to_cell( $s1f2_row, $col );
			$worksheet->write( $a, $col, $sum, $bodyNum );
			
			if ($col eq 8 or $col eq 13 or $col eq 18){
					if ($wtd_net_ly eq 0 or $mtd_net_ly eq 0 or $qtd_net_ly eq 0) {
						$worksheet->write( $a, $col+1, "", $bodyPct );}
					else{
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
						$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );}
			}
			elsif ($col eq 10 or $col eq 15 or $col eq 20){
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
			}
		}
		else{
			my $sum = '=SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . ')';
				$worksheet->write( $a, $col, $sum, $bodyNum );	
				
			if ($col eq 8 or $col eq 13 or $col eq 18){
					if ($wtd_net_ly eq 0 or $mtd_net_ly eq 0 or $qtd_net_ly eq 0) {
						$worksheet->write( $a, $col+1, "", $bodyPct );}
					else{
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
						$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );}
			}
			elsif ($col eq 10 or $col eq 15 or $col eq 20){
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
			}
		}		
	}
}
}


sub generate_data {

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh_rms = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh, ">", "vendor_export_sales.csv" or die "vendor_export_sales.csv: $!";

$test = qq{ 
SELECT 
CASE WHEN LOCATION IN ('3009','3012','6004','6009','6012') THEN 1 ELSE 0 END AS NEW_FLG, 
CASE WHEN LOCATION IN ('3009','3012','6004','6009','6012') THEN 0 ELSE 1 END AS MATURED_FLG,
LOCATION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, ATTRIB1 GROUP_CODE, ATTRIB2 GROUP_DESC, DIVISION, DIV_NAME DIVISION_DESC, GROUP_NO, GROUP_NAME, SUPPLIER, REPLACE(SUP_NAME, ',', ' ') AS SUP_NAME, COUNTRY, IMP_TYPE, 
SUM(NVL(TOTAL_RETAIL_MTD,0)) TOTAL_RETAIL_MTD, 
SUM(NVL(TOTAL_RETAIL_LY_MTD,0)) TOTAL_RETAIL_LY_MTD, 
0 AS TOTAL_RETAIL_QTD, 
0 AS TOTAL_RETAIL_LY_QTD
FROM
(
SELECT TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35) GROUP_NAME, SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49) SUP_NAME, CTY.UDA_VALUE_DESC COUNTRY, TYP.IMP_TYPE, SUM(TDH.TOTAL_RETAIL) TOTAL_RETAIL_MTD, 0 AS TOTAL_RETAIL_LY_MTD, 0 AS TOTAL_RETAIL_QTD, 0 AS TOTAL_RETAIL_LY_QTD
FROM TRAN_DATA_HISTORY TDH
  LEFT JOIN ITEM_MASTER MST ON TDH.ITEM = MST.ITEM
  LEFT JOIN DEPS ON MST.DEPT = DEPS.DEPT
  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
  LEFT JOIN DIVISION DIV ON GROUPS.DIVISION = DIV.DIVISION
  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = DIV.DIVISION
  LEFT JOIN ITEM_SUPPLIER SUP ON MST.ITEM = SUP.ITEM AND SUP.PRIMARY_SUPP_IND = 'Y'
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1204)CTY ON TDH.ITEM = CTY.ITEM
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC AS IMP_TYPE
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1103)TYP ON TDH.ITEM = TYP.ITEM
  LEFT JOIN SUPS ON SUP.SUPPLIER = SUPS.SUPPLIER
WHERE TDH.TRAN_CODE = 1 AND trunc(TDH.TRAN_DATE) BETWEEN '$mo_st_date_fld' AND '$mo_en_date_fld' AND (BI.MERCH_GROUP_CODE = 'SU' OR (BI.DIVISION = 8000 AND DEPS.DEPT = 8040))
GROUP BY TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35), SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49), CTY.UDA_VALUE_DESC, TYP.IMP_TYPE 

UNION ALL

SELECT TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35) GROUP_NAME, SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49) SUP_NAME, CTY.UDA_VALUE_DESC COUNTRY, TYP.IMP_TYPE, 0 AS TOTAL_RETAIL_MTD, SUM(TDH.TOTAL_RETAIL) AS TOTAL_RETAIL_LY_MTD, 0 AS TOTAL_RETAIL_QTD, 0 AS TOTAL_RETAIL_LY_QTD
FROM TRAN_DATA_HISTORY TDH
  LEFT JOIN ITEM_MASTER MST ON TDH.ITEM = MST.ITEM
  LEFT JOIN DEPS ON MST.DEPT = DEPS.DEPT
  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
  LEFT JOIN DIVISION DIV ON GROUPS.DIVISION = DIV.DIVISION
  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = DIV.DIVISION
  LEFT JOIN ITEM_SUPPLIER SUP ON MST.ITEM = SUP.ITEM AND SUP.PRIMARY_SUPP_IND = 'Y'
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1204)CTY ON TDH.ITEM = CTY.ITEM
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC AS IMP_TYPE
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1103)TYP ON TDH.ITEM = TYP.ITEM
  LEFT JOIN SUPS ON SUP.SUPPLIER = SUPS.SUPPLIER
WHERE TDH.TRAN_CODE = 1 AND trunc(TDH.TRAN_DATE) BETWEEN '$mo_st_date_fld_ly' AND 'mo_en_date_fld_ly' AND (BI.MERCH_GROUP_CODE = 'SU' OR (BI.DIVISION = 8000 AND DEPS.DEPT = 8040))
GROUP BY TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35), SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49), CTY.UDA_VALUE_DESC, TYP.IMP_TYPE 
)
GROUP BY 
CASE WHEN LOCATION IN ('3009','3012','6004','6009','6012') THEN 1 ELSE 0 END, 
CASE WHEN LOCATION IN ('3009','3012','6004','6009','6012') THEN 0 ELSE 1 END,
LOCATION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, ATTRIB1, ATTRIB2, DIVISION, DIV_NAME, GROUP_NO, GROUP_NAME, SUPPLIER, REPLACE(SUP_NAME, ',', ' '), COUNTRY, IMP_TYPE 
};

$testXX = qq{ 
SELECT 
CASE WHEN LOCATION IN ('3009','3012','6004','6009','6012') THEN 1 ELSE 0 END AS NEW_FLG, 
CASE WHEN LOCATION IN ('3009','3012','6004','6009','6012') THEN 0 ELSE 1 END AS MATURED_FLG,
replace(LOCATION, ',', ' ') LOCATION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, ATTRIB1 GROUP_CODE, ATTRIB2 GROUP_DESC, DIVISION, DIV_NAME DIVISION_DESC, GROUP_NO, GROUP_NAME, SUPPLIER, SUP_NAME, COUNTRY, IMP_TYPE, 
SUM(NVL(TOTAL_RETAIL_MTD,0)) TOTAL_RETAIL_MTD, 
SUM(NVL(TOTAL_RETAIL_LY_MTD,0)) TOTAL_RETAIL_LY_MTD, 
SUM(NVL(TOTAL_RETAIL_QTD,0)) TOTAL_RETAIL_QTD, 
SUM(NVL(TOTAL_RETAIL_LY_QTD,0)) TOTAL_RETAIL_LY_QTD
FROM
(
SELECT TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35) GROUP_NAME, SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49) SUP_NAME, CTY.UDA_VALUE_DESC COUNTRY, TYP.UDA_VALUE_DESC IMP_TYPE, SUM(TDH.TOTAL_RETAIL) TOTAL_RETAIL_MTD, 0 AS TOTAL_RETAIL_LY_MTD, 0 AS TOTAL_RETAIL_QTD, 0 AS TOTAL_RETAIL_LY_QTD
FROM TRAN_DATA_HISTORY TDH
  LEFT JOIN ITEM_MASTER MST ON TDH.ITEM = MST.ITEM
  LEFT JOIN DEPS ON MST.DEPT = DEPS.DEPT
  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
  LEFT JOIN DIVISION DIV ON GROUPS.DIVISION = DIV.DIVISION
  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = DIV.DIVISION
  LEFT JOIN ITEM_SUPPLIER SUP ON MST.ITEM = SUP.ITEM AND SUP.PRIMARY_SUPP_IND = 'Y'
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1204)CTY ON TDH.ITEM = CTY.ITEM
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1103)TYP ON TDH.ITEM = TYP.ITEM
  LEFT JOIN SUPS ON SUP.SUPPLIER = SUPS.SUPPLIER
WHERE TDH.TRAN_CODE = 1 AND trunc(TDH.TRAN_DATE) BETWEEN '$mo_st_date_fld' AND '$mo_en_date_fld' AND (BI.MERCH_GROUP_CODE = 'SU' OR (BI.DIVISION = 8000 AND DEPS.DEPT = 8040))
GROUP BY TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35), SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49), CTY.UDA_VALUE_DESC, TYP.UDA_VALUE_DESC 

UNION ALL

SELECT TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35) GROUP_NAME, SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49) SUP_NAME, CTY.UDA_VALUE_DESC COUNTRY, TYP.UDA_VALUE_DESC IMP_TYPE, 0 AS TOTAL_RETAIL_MTD, SUM(TDH.TOTAL_RETAIL) AS TOTAL_RETAIL_LY_MTD, 0 AS TOTAL_RETAIL_QTD, 0 AS TOTAL_RETAIL_LY_QTD
FROM TRAN_DATA_HISTORY TDH
  LEFT JOIN ITEM_MASTER MST ON TDH.ITEM = MST.ITEM
  LEFT JOIN DEPS ON MST.DEPT = DEPS.DEPT
  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
  LEFT JOIN DIVISION DIV ON GROUPS.DIVISION = DIV.DIVISION
  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = DIV.DIVISION
  LEFT JOIN ITEM_SUPPLIER SUP ON MST.ITEM = SUP.ITEM AND SUP.PRIMARY_SUPP_IND = 'Y'
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1204)CTY ON TDH.ITEM = CTY.ITEM
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1103)TYP ON TDH.ITEM = TYP.ITEM
  LEFT JOIN SUPS ON SUP.SUPPLIER = SUPS.SUPPLIER
WHERE TDH.TRAN_CODE = 1 AND trunc(TDH.TRAN_DATE) BETWEEN '$mo_st_date_fld_ly' AND 'mo_en_date_fld_ly' AND (BI.MERCH_GROUP_CODE = 'SU' OR (BI.DIVISION = 8000 AND DEPS.DEPT = 8040))
GROUP BY TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35), SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49), CTY.UDA_VALUE_DESC, TYP.UDA_VALUE_DESC 

UNION ALL

SELECT TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35) GROUP_NAME, SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49) SUP_NAME, CTY.UDA_VALUE_DESC COUNTRY, TYP.UDA_VALUE_DESC IMP_TYPE, 0 AS TOTAL_RETAIL_MTD, 0 AS TOTAL_RETAIL_LY_MTD, SUM(TDH.TOTAL_RETAIL) AS TOTAL_RETAIL_QTD, 0 AS TOTAL_RETAIL_LY_QTD
FROM TRAN_DATA_HISTORY TDH
  LEFT JOIN ITEM_MASTER MST ON TDH.ITEM = MST.ITEM
  LEFT JOIN DEPS ON MST.DEPT = DEPS.DEPT
  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
  LEFT JOIN DIVISION DIV ON GROUPS.DIVISION = DIV.DIVISION
  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = DIV.DIVISION
  LEFT JOIN ITEM_SUPPLIER SUP ON MST.ITEM = SUP.ITEM AND SUP.PRIMARY_SUPP_IND = 'Y'
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1204)CTY ON TDH.ITEM = CTY.ITEM
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1103)TYP ON TDH.ITEM = TYP.ITEM
  LEFT JOIN SUPS ON SUP.SUPPLIER = SUPS.SUPPLIER
WHERE TDH.TRAN_CODE = 1 AND trunc(TDH.TRAN_DATE) BETWEEN '$qu_st_date_fld' AND '$mo_en_date_fld' AND (BI.MERCH_GROUP_CODE = 'SU' OR (BI.DIVISION = 8000 AND DEPS.DEPT = 8040))
GROUP BY TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35), SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49), CTY.UDA_VALUE_DESC, TYP.UDA_VALUE_DESC 

UNION ALL

SELECT TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35) GROUP_NAME, SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49) SUP_NAME, CTY.UDA_VALUE_DESC COUNTRY, TYP.UDA_VALUE_DESC IMP_TYPE, 0 AS TOTAL_RETAIL_MTD, 0 AS TOTAL_RETAIL_LY_MTD, 0 AS TOTAL_RETAIL_QTD, SUM(TDH.TOTAL_RETAIL) AS TOTAL_RETAIL_LY_QTD
FROM TRAN_DATA_HISTORY TDH
  LEFT JOIN ITEM_MASTER MST ON TDH.ITEM = MST.ITEM
  LEFT JOIN DEPS ON MST.DEPT = DEPS.DEPT
  LEFT JOIN GROUPS ON DEPS.GROUP_NO = GROUPS.GROUP_NO
  LEFT JOIN DIVISION DIV ON GROUPS.DIVISION = DIV.DIVISION
  LEFT JOIN BI_MERCH_GROUP BI ON BI.DIVISION = DIV.DIVISION
  LEFT JOIN ITEM_SUPPLIER SUP ON MST.ITEM = SUP.ITEM AND SUP.PRIMARY_SUPP_IND = 'Y'
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1204)CTY ON TDH.ITEM = CTY.ITEM
  LEFT JOIN (SELECT DISTINCT MST.ITEM, VAL.UDA_VALUE_DESC 
			FROM ITEM_MASTER MST 
			  LEFT JOIN UDA_ITEM_LOV LOV ON MST.ITEM = LOV.ITEM
			  LEFT JOIN UDA_VALUES VAL ON LOV.UDA_ID = VAL.UDA_ID AND LOV.UDA_VALUE = VAL.UDA_VALUE
			WHERE   LOV.UDA_ID = 1103)TYP ON TDH.ITEM = TYP.ITEM
  LEFT JOIN SUPS ON SUP.SUPPLIER = SUPS.SUPPLIER
WHERE TDH.TRAN_CODE = 1 AND trunc(TDH.TRAN_DATE) BETWEEN '$qu_st_date_fld_ly' AND '$mo_en_date_fld_ly' AND (BI.MERCH_GROUP_CODE = 'SU' OR (BI.DIVISION = 8000 AND DEPS.DEPT = 8040))
GROUP BY TDH.LOCATION, BI.MERCH_GROUP_CODE, BI.MERCH_GROUP_DESC, BI.ATTRIB1, BI.ATTRIB2, BI.DIVISION, DIV.DIV_NAME, DEPS.GROUP_NO, SUBSTR(GROUPS.GROUP_NAME,1,35), SUP.SUPPLIER, SUBSTR(SUPS.SUP_NAME,1,49), CTY.UDA_VALUE_DESC, TYP.UDA_VALUE_DESC 
)
GROUP BY 
CASE WHEN LOCATION IN ('3009','3012','6004','6009','6012') THEN 1 ELSE 0 END, 
CASE WHEN LOCATION IN ('3009','3012','6004','6009','6012') THEN 0 ELSE 1 END,
replace(LOCATION, ',', ' '), MERCH_GROUP_CODE, MERCH_GROUP_DESC, ATTRIB1, ATTRIB2, DIVISION, DIV_NAME, GROUP_NO, GROUP_NAME, SUPPLIER, SUP_NAME, COUNTRY, IMP_TYPE 
};

my $sth = $dbh_rms->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_lc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "vendor_export_sales.csv: $!";
 
$sth->finish();
$dbh_rms->disconnect;

}

sub insert_data {

my $truncate = $dbh->prepare( qq{ 
TRUNCATE TABLE METRO_IT_VENDOR_IMPORTS4 
});
$truncate->execute();

$truncate->finish();

print "Done truncating METRO_IT_VENDOR_IMPORTS4... \nPreparing to Insert... \n";

my $sth_insert = $dbh->prepare( q{
INSERT INTO METRO_IT_VENDOR_IMPORTS4 (NEW_FLG, MATURED_FLG, LOCATION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC, GROUP_NO, GROUP_NAME, SUPPLIER, SUP_NAME, COUNTRY, TOTAL_RETAIL_MTD, TOTAL_RETAIL_LY_MTD, TOTAL_RETAIL_QTD, TOTAL_RETAIL_LY_QTD)
VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? ) 
});
  
open FH1, "<vendor_export_sales.csv" or die "Unable to open vendor_export_sales.csv: $!";
while (<FH1>) {
	chomp;
    my ( $new_flg, $matured_flg, $location, $merch_group_code, $merch_group_desc, $group_code, $group_desc, $division, $division_desc, $group_no, $group_name, $supplier, $sup_name, $country, $total_retail_mtd, $total_retail_ly_mtd, $total_retail_qtd, $total_retail_ly_qtd ) = split (/,/);
	
	$sth_insert->execute( $new_flg, $matured_flg, $location, $merch_group_code, $merch_group_desc, $group_code, $group_desc, $division, $division_desc, $group_no, $group_name, $supplier, $sup_name, $country, $total_retail_mtd, $total_retail_ly_mtd, $total_retail_qtd, $total_retail_ly_qtd );	
}
close FH1;

$dbh->commit;
$sth_insert->finish();
$dbh->disconnect;

}

# mailer
sub mail_grp1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1, $attachment_file_2 ) = @ARGV;

$to = ' arthur.emmanuel@metrogaisano.com, frank.gaisano@metrogaisano.com, gerry.guanlao@metrogaisano.com, eric.redona@metrogaisano.com, lucille.malazarte@metrogaisano.com, tricia.luntao@metrogaisano.com, jj.moreno@metrogaisano.com, rose.jose@metrogaisano.com, rex.cabanilla@metrogaisano.com  ';

$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, frank.naquines@metrogaisano.com, cham.burgos@metrogaisano.com ';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';

$subject = 'Daily Sales Performance as of ' . $as_of;

$msgbody_file = 'message_BI.txt';

$attachment_file_1 = "Daily Sales Performance - Summary (as of $as_of) v1.6.xlsx";
$attachment_file_2 = "Daily Sales Performance - Summary (as of $as_of) v1.6.pdf";

my $msgbody = read_file( $msgbody_file );

my $attachment_data_1 = encode_base64( read_file( $attachment_file_1, 1 ));
my $attachment_data_2 = encode_base64( read_file( $attachment_file_2, 1 ));

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
Content-Type: application/octet-stream; name="$attachment_file_1"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$attachment_file_1"
$attachment_data_1
$boundary
Content-Type: application/octet-stream; name="$attachment_file_2"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$attachment_file_2"
$attachment_data_2
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







