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
use HTML::Entities;
use HTML::Table::FromDatabase;
use Date::Calc qw( Today Add_Delta_Days Month_to_Text);

### updated chronology of function calls, NB stores changed to Supermarket

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
	SELECT WEEK_NUMBER_THIS_YEAR, DATE_KEY, DATE_FLD, WEEK_ST_DATE_KEY, WEEK_END_DATE_KEY, YTD_DATE_KEY, YTD_DATE_FLD, YTD_DATE_KEY_LY, YTD_DATE_FLD_LY  FROM
	  (SELECT WEEK_NUMBER_THIS_YEAR, DATE_KEY, TO_CHAR(DATE_FLD, 'DD Mon YYYY') DATE_FLD, WEEK_ST_DATE_KEY, WEEK_END_DATE_KEY 
		FROM DIM_DATE
		WHERE DATE_FLD = (SELECT AGG_DLY_END_DATE_FLD FROM ADMIN_ETL_DATE_PARAMETER)),
	  (SELECT DATE_KEY YTD_DATE_KEY, TO_CHAR(DATE_FLD, 'DD Mon YYYY') YTD_DATE_FLD, DATE_KEY_LY YTD_DATE_KEY_LY, TO_CHAR(DATE_FLD_LY, 'DD Mon YYYY') YTD_DATE_FLD_LY
	  FROM DIM_DATE_PRL WHERE DAY_IN_YEAR = 1 AND QUARTER = 1 AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') AS YEAR FROM DUAL))
	 };

	my $sth_date_1 = $dbh->prepare ($date);
	 $sth_date_1->execute;

	while (my $x = $sth_date_1->fetchrow_hashref()) {
		$wk_st_date_key = $x->{WEEK_ST_DATE_KEY};
		$wk_en_date_key = $x->{DATE_KEY};
		$wk_number = $x->{WEEK_NUMBER_THIS_YEAR};
		$as_of = $x->{DATE_FLD};
		$yr_st_date_key = $x->{YTD_DATE_KEY};
		$yr_st_date_fld = $x->{YTD_DATE_FLD};
		$yr_st_date_key_ly = $x->{YTD_DATE_KEY_LY};
		$yr_st_date_fld_ly = $x->{YTD_DATE_FLD_LY};
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
	
	$date4 = qq{ 
	SELECT WTD_ST_DT, WTD_ST_DT_LY, WTD_EN_DT, WTD_EN_DT_LY FROM
		(SELECT TO_CHAR(DATE_FLD, 'DD Mon YYYY') AS WTD_ST_DT, TO_CHAR(DATE_FLD_LY, 'DD Mon YYYY') AS WTD_ST_DT_LY 
			FROM DIM_DATE WHERE DATE_KEY = (SELECT WEEK_ST_DATE_KEY FROM DIM_DATE WHERE TRUNC(DATE_FLD) = TRUNC((SELECT DISTINCT UPDATE_DATE FROM METRO_IT_SALES_DEPT_WTD)))),
		(SELECT TO_CHAR(DATE_FLD-1, 'DD Mon YYYY') AS WTD_EN_DT, TO_CHAR(DATE_FLD_LY-1, 'DD Mon YYYY') AS WTD_EN_DT_LY 
			FROM DIM_DATE WHERE TRUNC(DATE_FLD) = TRUNC((SELECT DISTINCT UPDATE_DATE FROM METRO_IT_SALES_DEPT_WTD)))
	 };

	my $sth_date_4 = $dbh->prepare ($date4);
	 $sth_date_4->execute;

	while (my $x = $sth_date_4->fetchrow_hashref()) {
		$wtd_st_dt = $x->{WTD_ST_DT};
		$wtd_st_dt_ly = $x->{WTD_ST_DT_LY};
		$wtd_en_dt = $x->{WTD_EN_DT};
		$wtd_en_dt_ly = $x->{WTD_EN_DT_LY};
	}
 
	$workbook = Excel::Writer::XLSX->new("Daily Sales Performance - Concession (as of $as_of) V1.4.xlsx");
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
	
	# &generate_csv;
	
	&new_sheet($sheet = "Summary");
	&call_str;

	&new_sheet($sheet = "GenMerch_Spmkt");
	&call_str_merchandise;
	
	&new_sheet_2($sheet = "Department");			
	&call_div;
		
	$workbook->close();
	
	&mail_grp1;
	&mail_grp2;
	
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
 
# function call
sub call_div {

$a = 11, $counter = 0;

$worksheet->write($a-11, 2, "Daily Sales Performance - Concession", $bold1);
$worksheet->write($a-10, 2, "WTD: $wtd_st_dt - $wtd_en_dt vs $wtd_st_dt_ly - $wtd_en_dt_ly");
$worksheet->write($a-9, 2, "MTD: $mo_st_date_fld - $mo_en_date_fld vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-8, 2, "QTD: $qu_st_date_fld - $mo_en_date_fld vs $qu_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-7, 2, "YTD: $yr_st_date_fld - $mo_en_date_fld vs $yr_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-6, 2, "As of $as_of");

##========================= COMP STORES ===========================##

&heading_2;
&heading;
&query_dept($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 1, $matured_flg2 = 1, $loc_desc = "COMP STORES");

##========================= ALL STORES ===========================##

$a += 6;
&heading_2;
&heading;
&query_dept($new_flg1 = 0, $new_flg2 = 1, $matured_flg1 = 0, $matured_flg2 = 1, $loc_desc = "ALL STORES");

##========================= BY STORE ===========================##

foreach my $i ( '2001', '2001W', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', '2013', '2223', '3001', '3002', '3003', '3004', '3005', '3006', '3007', '3009', '3010', '3012', '4003', '4004', '6001', '6002', '6003', '6004', '6005', '6009', '6010', '6012' ){ 
#foreach my $i ( '2001', '2001W' ){ 
	$a += 6;	
	&heading_2;
	&heading;
	&query_dept_store($store = $i);

}

}

sub call_str {

$a=11, $counter=0;

$worksheet->write($a-11, 3, "Daily Sales Performance - Concession", $bold1);
$worksheet->write($a-10, 3, "WTD: $wtd_st_dt - $wtd_en_dt vs $wtd_st_dt_ly - $wtd_en_dt_ly");
$worksheet->write($a-9, 3, "MTD: $mo_st_date_fld - $mo_en_date_fld vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-8, 3, "QTD: $qu_st_date_fld - $mo_en_date_fld vs $qu_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-7, 3, "YTD: $yr_st_date_fld - $mo_en_date_fld vs $yr_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-6, 3, "As of $as_of");

$worksheet->write($a-4, 3, "Summary", $bold);

&heading;

$worksheet->merge_range( $a-2, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a-1, 6, 'Format', $subhead );

&query_summary($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 1, $matured_flg2 = 1, $summary_label = 'COMP');
&query_summary($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 0, $matured_flg2 = 0, $summary_label = 'NEW');
&query_summary($new_flg1 = 0, $new_flg2 = 1, $matured_flg1 = 0, $matured_flg2 = 1, $summary_label = 'ALL');

$a+=5; 

$worksheet->write($a-4, 3, "Per Store", $bold);

&heading;

$worksheet->merge_range( $a-2, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a-1, 4, 'Type', $subhead );
$worksheet->merge_range( $a-2, 5, $a-1, 5, 'Code', $subhead );
$worksheet->merge_range( $a-2, 6, $a-1, 6, 'Desc', $subhead );

&query_summary_by_store($store_format = 'DEPARTMENT STORE');
&query_summary_by_store($store_format = 'SUPERMARKET');
&query_summary_by_store($store_format = 'HYPERMARKET');

}

sub call_str_merchandise {

$a=11, $counter=0;

$worksheet->write($a-11, 3, "Daily Sales Performance - Concession", $bold1);
$worksheet->write($a-10, 3, "WTD: $wtd_st_dt - $wtd_en_dt vs $wtd_st_dt_ly - $wtd_en_dt_ly");
$worksheet->write($a-9, 3, "MTD: $mo_st_date_fld - $mo_en_date_fld vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-8, 3, "QTD: $qu_st_date_fld - $mo_en_date_fld vs $qu_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-7, 3, "YTD: $yr_st_date_fld - $mo_en_date_fld vs $yr_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-6, 3, "As of $as_of");

$worksheet->write($a-4, 3, "Summary", $bold);

&heading_3;

$worksheet->merge_range( $a-2, 3, $a, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a, 6, 'Format', $subhead );

$a += 1;

&query_summary_merchandise($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 1, $matured_flg2 = 1, $summary_label = 'COMP');
&query_summary_merchandise($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 0, $matured_flg2 = 0, $summary_label = 'NEW');
&query_summary_merchandise($new_flg1 = 0, $new_flg2 = 1, $matured_flg1 = 0, $matured_flg2 = 1, $summary_label = 'ALL');

$a+=5; 

$worksheet->write($a-4, 3, "Per Store", $bold);

&heading_3;;

$worksheet->merge_range( $a-2, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a-1, 4, 'Type', $subhead );
$worksheet->merge_range( $a-2, 5, $a-1, 5, 'Code', $subhead );
$worksheet->merge_range( $a-2, 6, $a-1, 6, 'Desc', $subhead );

&query_summary_merchandise_by_store($store_format = 'DEPARTMENT STORE');
&query_summary_merchandise_by_store($store_format = 'SUPERMARKET');
&query_summary_merchandise_by_store($store_format = 'HYPERMARKET');

}

# create sheet
sub new_sheet{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(92);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
#$worksheet->set_print_scale( 100 );
$worksheet->fit_to_pages( 1, 1 );
$worksheet->set_margins( 0.001 );
$worksheet->conditional_formatting( 'F9:AU6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );

$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 27 );

$worksheet->set_column( 7, 8, 8 );
$worksheet->set_column( 9, 9, 7 );
# $worksheet->set_column( 10, 10, 8 );
# $worksheet->set_column( 11, 11, 7 );
$worksheet->set_column( 10, 11, undef, undef, 1 );
$worksheet->set_column( 12, 13, 10 );
$worksheet->set_column( 14, 14, 7 );
# $worksheet->set_column( 15, 15, 10 );
# $worksheet->set_column( 16, 16, 7 );
$worksheet->set_column( 15, 16, undef, undef, 1 );

$worksheet->set_column( 17, 18, 10 );
$worksheet->set_column( 19, 19, 7 );
# $worksheet->set_column( 20, 20, 10 );
# $worksheet->set_column( 21, 21, 7 );
$worksheet->set_column( 20, 21, undef, undef, 1 );
$worksheet->set_column( 22, 23, 10 );
$worksheet->set_column( 24, 24, 7 );
# $worksheet->set_column( 25, 25, 10 );
# $worksheet->set_column( 26, 26, 7 );
$worksheet->set_column( 25, 26, undef, undef, 1 );

$worksheet->set_column( 27, 28, 10 );
$worksheet->set_column( 29, 29, 7 );
# $worksheet->set_column( 30, 30, 10 );
# $worksheet->set_column( 31, 31, 7 );
$worksheet->set_column( 30, 31, undef, undef, 1 );
$worksheet->set_column( 32, 33, 10 );
$worksheet->set_column( 34, 34, 7 );
# $worksheet->set_column( 35, 35, 10 );
# $worksheet->set_column( 36, 36, 7 );
$worksheet->set_column( 35, 36, undef, undef, 1 );

$worksheet->set_column( 37, 38, 10 );
$worksheet->set_column( 39, 39, 7 );
# $worksheet->set_column( 40, 40, 10 );
# $worksheet->set_column( 41, 41, 7 );
$worksheet->set_column( 40, 41, undef, undef, 1 );
$worksheet->set_column( 42, 43, 10 );
$worksheet->set_column( 44, 44, 7 );
# $worksheet->set_column( 45, 45, 10 );
# $worksheet->set_column( 46, 46, 7 );
$worksheet->set_column( 45, 46, undef, undef, 1 );


}

sub new_sheet_2{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom( 92 );
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
#$worksheet->set_print_scale( 100 );
$worksheet->fit_to_pages( 1, 1 );
$worksheet->set_margins( 0.001 );
$worksheet->conditional_formatting( 'F9:AU6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );

$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 27 );

$worksheet->set_column( 7, 8, 8 );
$worksheet->set_column( 9, 9, 7 );
# $worksheet->set_column( 10, 10, 8 );
# $worksheet->set_column( 11, 11, 7 );
$worksheet->set_column( 10, 11, undef, undef, 1 );
$worksheet->set_column( 12, 13, 10 );
$worksheet->set_column( 14, 14, 7 );
# $worksheet->set_column( 15, 15, 10 );
# $worksheet->set_column( 16, 16, 7 );
$worksheet->set_column( 15, 16, undef, undef, 1 );

$worksheet->set_column( 17, 18, 10 );
$worksheet->set_column( 19, 19, 7 );
# $worksheet->set_column( 20, 20, 10 );
# $worksheet->set_column( 21, 21, 7 );
$worksheet->set_column( 20, 21, undef, undef, 1 );
$worksheet->set_column( 22, 23, 10 );
$worksheet->set_column( 24, 24, 7 );
# $worksheet->set_column( 25, 25, 10 );
# $worksheet->set_column( 26, 26, 7 );
$worksheet->set_column( 25, 26, undef, undef, 1 );

$worksheet->set_column( 27, 28, 10 );
$worksheet->set_column( 29, 29, 7 );
# $worksheet->set_column( 30, 30, 10 );
# $worksheet->set_column( 31, 31, 7 );
$worksheet->set_column( 30, 31, undef, undef, 1 );
$worksheet->set_column( 32, 33, 10 );
$worksheet->set_column( 34, 34, 7 );
# $worksheet->set_column( 35, 35, 10 );
# $worksheet->set_column( 36, 36, 7 );
$worksheet->set_column( 35, 36, undef, undef, 1 );

$worksheet->set_column( 37, 38, 10 );
$worksheet->set_column( 39, 39, 7 );
# $worksheet->set_column( 40, 40, 10 );
# $worksheet->set_column( 41, 41, 7 );
$worksheet->set_column( 40, 41, undef, undef, 1 );
$worksheet->set_column( 42, 43, 10 );
$worksheet->set_column( 44, 44, 7 );
# $worksheet->set_column( 45, 45, 10 );
# $worksheet->set_column( 46, 46, 7 );
$worksheet->set_column( 45, 46, undef, undef, 1 );

}

# headers
sub heading {

$worksheet->write($a-3, 3, "in 000's", $script);
$worksheet->merge_range( $a-2, 7, $a-2, 11, 'WTD', $subhead );
$worksheet->merge_range( $a-2, 12, $a-2, 16, 'MTD', $subhead );
$worksheet->merge_range( $a-2, 17, $a-2, 21, 'QTD', $subhead );
$worksheet->merge_range( $a-2, 22, $a-2, 26, 'YTD', $subhead );

foreach my $i ( 7, 12, 17, 22 ) {
	$worksheet->write($a-1, $i, "TY", $subhead);
	$worksheet->write($a-1, $i+1, "LY", $subhead);
	$worksheet->write($a-1, $i+2, "Growth", $subhead);
	$worksheet->write($a-1, $i+3, "Budget", $subhead);
	$worksheet->write($a-1, $i+4, "vs Budget", $subhead);
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

sub heading_3 {

$worksheet->write($a-3, 3, "in 000's", $script);
$worksheet->merge_range( $a-2, 7, $a-2, 16, 'WTD', $subhead );
$worksheet->merge_range( $a-2, 17, $a-2, 26, 'MTD', $subhead );
$worksheet->merge_range( $a-2, 27, $a-2, 36, 'QTD', $subhead );
$worksheet->merge_range( $a-2, 37, $a-2, 46, 'YTD', $subhead );

$worksheet->merge_range( $a-1, 7, $a-1, 11, 'General Merchandise', $subhead );
$worksheet->merge_range( $a-1, 12, $a-1, 16, 'Supermarket', $subhead );
$worksheet->merge_range( $a-1, 17, $a-1, 21, 'General Merchandise', $subhead );
$worksheet->merge_range( $a-1, 22, $a-1, 26, 'Supermarket', $subhead );
$worksheet->merge_range( $a-1, 27, $a-1, 31, 'General Merchandise', $subhead );
$worksheet->merge_range( $a-1, 32, $a-1, 36, 'Supermarket', $subhead );
$worksheet->merge_range( $a-1, 37, $a-1, 41, 'General Merchandise', $subhead );
$worksheet->merge_range( $a-1, 42, $a-1, 46, 'Supermarket', $subhead );

foreach my $i ( 7, 12, 17, 22, 27, 32, 37, 42 ) {
	$worksheet->write($a, $i, "TY", $subhead);
	$worksheet->write($a, $i+1, "LY", $subhead);
	$worksheet->write($a, $i+2, "Growth", $subhead);
	$worksheet->write($a, $i+3, "Budget", $subhead);
	$worksheet->write($a, $i+4, "vs Budget", $subhead);
}

}

#sheet 1
sub query_summary{

$sls = $dbh->prepare (qq{
	SELECT 
		SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
		SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
		SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
		SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
	FROM
		(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A LEFT JOIN
		(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, 
						SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
		FROM METRO_IT_SALES_DEPT_WTD
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) AND MERCH_GROUP_CODE <> 'Z_OT'
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END
		UNION ALL
		SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
						FROM METRO_IT_SALES_DEPT 
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END
		UNION ALL
		SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
						FROM METRO_IT_SALES_DEPT 
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' 
			AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END
		UNION ALL
		SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						SUM(TARGET_SALE_VAL) YTD_TARGET, SUM(NET_SALE_CON_TY) YTD_NET_TY, SUM(NET_SALE_CON_LY) YTD_NET_LY 
						FROM METRO_IT_SALES_DEPT 
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' 
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)B ON A.FORMAT = B.FORMAT
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
		SELECT A.FORMAT,
			SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
			SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
			SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
			SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
		FROM
			(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A LEFT JOIN
			(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, 
						SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
			FROM METRO_IT_SALES_DEPT_WTD
			WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) AND MERCH_GROUP_CODE <> 'Z_OT'
			GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END
			UNION ALL
			SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
							FROM METRO_IT_SALES_DEPT 
			WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
			GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END
			UNION ALL
			SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY, 
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
							FROM METRO_IT_SALES_DEPT 
			WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' 
				AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
			GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END
			UNION ALL
			SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
							SUM(TARGET_SALE_VAL) YTD_TARGET, SUM(NET_SALE_CON_TY) YTD_NET_TY, SUM(NET_SALE_CON_LY) YTD_NET_LY 
							FROM METRO_IT_SALES_DEPT 
			WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
				AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' 
			GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)B ON A.FORMAT=B.FORMAT
		GROUP BY A.FORMAT
		ORDER BY 1
		});
	$sls1->execute();
						
		while(my $s = $sls1->fetchrow_hashref()){
								
		$worksheet->merge_range( $a, 4, $a, 6, $s->{FORMAT}, $desc );
		$worksheet->write($a,7, $s->{WTD_NET_TY},$border1);
		$worksheet->write($a,8, $s->{WTD_NET_LY},$border1);
		$worksheet->write($a,10, $s->{WTD_TARGET},$border1);
			if ($s->{WTD_NET_LY} <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$subt); }
									
			if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
				$worksheet->write($a,11, "",$subt); }
			else{
				$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$subt); }
									
		$worksheet->write($a,12, $s->{MTD_NET_TY},$border1);
		$worksheet->write($a,13, $s->{MTD_NET_LY},$border1);
		$worksheet->write($a,15, $s->{MTD_TARGET},$border1);
			if ($s->{MTD_NET_LY} <= 0){
				$worksheet->write($a,14, "",$subt); }
			else{
				$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$subt); }
									
			if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$subt); }
										
		$worksheet->write($a,17, $s->{QTD_NET_TY},$border1);
		$worksheet->write($a,18, $s->{QTD_NET_LY},$border1);
		$worksheet->write($a,20, $s->{QTD_TARGET},$border1);
			if ($s->{QTD_NET_LY} <= 0){
				$worksheet->write($a,19, "",$subt); }
			else{
				$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$subt); }
									
			if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
				$worksheet->write($a,21, "",$subt); }
			else{
				$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$subt); }
										
		$worksheet->write($a,22, $s->{YTD_NET_TY},$border1);
		$worksheet->write($a,23, $s->{YTD_NET_LY},$border1);
		$worksheet->write($a,25, $s->{YTD_TARGET},$border1);
			if ($s->{YTD_NET_LY} <= 0){
				$worksheet->write($a,24, "",$subt); }
			else{
				$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$subt); }
									
			if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
				$worksheet->write($a,26, "",$subt); }
			else{
				$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$subt); }
												
		$a++;
		$counter++;
					
	}
	
	$worksheet->write($a,7, $s->{WTD_NET_TY},$bodyNum);
	$worksheet->write($a,8, $s->{WTD_NET_LY},$bodyNum);
	$worksheet->write($a,10, $s->{WTD_TARGET},$bodyNum);
		if ($s->{WTD_NET_LY} <= 0){
			$worksheet->write($a,9, "",$bodyPct); }
		else{
			$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$bodyPct); }
						
		if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
			$worksheet->write($a,11, "",$bodyPct); }
		else{
			$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$bodyPct); }
						
	$worksheet->write($a,12, $s->{MTD_NET_TY},$bodyNum);
	$worksheet->write($a,13, $s->{MTD_NET_LY},$bodyNum);
	$worksheet->write($a,15, $s->{MTD_TARGET},$bodyNum);
		if ($s->{MTD_NET_LY} <= 0){
			$worksheet->write($a,14, "",$bodyPct); }
		else{
			$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$bodyPct); }
							
		if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
			$worksheet->write($a,16, "",$bodyPct); }
		else{
			$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$bodyPct); }
						
	$worksheet->write($a,17, $s->{QTD_NET_TY},$bodyNum);
	$worksheet->write($a,18, $s->{QTD_NET_LY},$bodyNum);
	$worksheet->write($a,20, $s->{QTD_TARGET},$bodyNum);
		if ($s->{QTD_NET_LY} <= 0){
			$worksheet->write($a,19, "",$bodyPct); }
		else{
			$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$bodyPct); }
						
		if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
			$worksheet->write($a,21, "",$bodyPct); }
		else{
			$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$bodyPct); }
							
	$worksheet->write($a,22, $s->{YTD_NET_TY},$bodyNum);
	$worksheet->write($a,23, $s->{YTD_NET_LY},$bodyNum);
	$worksheet->write($a,25, $s->{YTD_TARGET},$bodyNum);
		if ($s->{YTD_NET_LY} <= 0){
			$worksheet->write($a,24, "",$bodyPct); }
		else{
			$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$bodyPct); }
						
		if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
			$worksheet->write($a,26, "",$bodyPct); }
		else{
			$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$bodyPct); }
	
	$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
	$worksheet->merge_range( $a-$counter, 3, $a, 3, $summary_label, $border2 );
	
	$counter = 0;
	$a++;
}

$sls->finish();
$sls1->finish();

}

sub query_summary_by_store{

$sls = $dbh->prepare (qq{
	SELECT 
		SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
		SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
		SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
		SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
	FROM
	  (SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
			  SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
	  FROM METRO_IT_SALES_DEPT_WTD
	  WHERE MERCH_GROUP_CODE <> 'Z_OT'
	  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  
	  UNION ALL	  
	  SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
			  0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
			  SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY, 
			  0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
			  0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
			  FROM METRO_IT_SALES_DEPT 
	  WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
	  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG	  
	  UNION ALL	  
	  SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
			  0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
			  0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
			  SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY, 
			  0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
			  FROM METRO_IT_SALES_DEPT 
	  WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' 
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
	  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG	  
	  UNION ALL	  
	  SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
			  0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
			  0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
			  0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
			  SUM(TARGET_SALE_VAL) YTD_TARGET, SUM(NET_SALE_CON_TY) YTD_NET_TY, SUM(NET_SALE_CON_LY) YTD_NET_LY
			  FROM METRO_IT_SALES_DEPT 
	  WHERE (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' 
	  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)
	WHERE FORMAT = '$store_format'
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
		SELECT MATURED_FLG, CASE WHEN MATURED_FLG = 1 THEN 'COMP' ELSE 'NEW' END AS FLG_DESC, 
			SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
			SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
			SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
			SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
		FROM
		  (	SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
			  SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
		  FROM METRO_IT_SALES_DEPT_WTD
		  WHERE MERCH_GROUP_CODE <> 'Z_OT'
		  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  
		  UNION ALL	  
		  SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
				  0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
				  SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY, 
				  0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
				  0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
				  FROM METRO_IT_SALES_DEPT 
		  WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
		  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG	  
		  UNION ALL	  
		  SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
				  0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
				  0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
				  SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY, 
				  0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
				  FROM METRO_IT_SALES_DEPT 
		  WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' 
			AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG	  
		  UNION ALL	  
		  SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
				  0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
				  0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
				  0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
				  SUM(TARGET_SALE_VAL) YTD_TARGET, SUM(NET_SALE_CON_TY) YTD_NET_TY, SUM(NET_SALE_CON_LY) YTD_NET_LY
				  FROM METRO_IT_SALES_DEPT 
		  WHERE (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' 
		  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)
		WHERE FORMAT = '$store_format'
		GROUP BY MATURED_FLG, CASE WHEN MATURED_FLG = 1 THEN 'COMP' ELSE 'NEW' END
		ORDER BY MATURED_FLG DESC
		});
	$sls1->execute();
		
		$format_counter = $a;
		while(my $s = $sls1->fetchrow_hashref()){
		$flg = $s->{MATURED_FLG};
		$flg_desc = $s->{FLG_DESC};

		$sls2 = $dbh->prepare (qq{
			SELECT STORE_CODE, STORE_DESCRIPTION, 
				SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
				SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
				SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
				SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
			FROM
			  (	SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
			  SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				  FROM METRO_IT_SALES_DEPT_WTD
				  WHERE MERCH_GROUP_CODE <> 'Z_OT' AND MATURED_FLG = '$flg'
				  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  
				  UNION ALL	  
				  SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
						  0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						  SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY, 
						  0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						  0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
						  FROM METRO_IT_SALES_DEPT 
				  WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT'  AND MATURED_FLG = '$flg' AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
				  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG	  
				  UNION ALL	  
				  SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
						  0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						  0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						  SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY, 
						  0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY 
						  FROM METRO_IT_SALES_DEPT 
				  WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' AND MATURED_FLG = '$flg' 
					AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
				  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG	  
				  UNION ALL	  
				  SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG,
						  0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						  0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						  0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						  SUM(TARGET_SALE_VAL) YTD_TARGET, SUM(NET_SALE_CON_TY) YTD_NET_TY, SUM(NET_SALE_CON_LY) YTD_NET_LY
						  FROM METRO_IT_SALES_DEPT 
				  WHERE (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
						AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE <> 'Z_OT' AND MATURED_FLG = '$flg' 
				  GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)
			WHERE FORMAT = '$store_format'
			GROUP BY STORE_CODE, STORE_DESCRIPTION
			ORDER BY 1
			});
		$sls2->execute();
			
			while(my $s = $sls2->fetchrow_hashref()){
									
			$worksheet->write( $a, 5, $s->{STORE_CODE}, $desc );
			$worksheet->write( $a, 6, $s->{STORE_DESCRIPTION}, $desc );
			$worksheet->write($a,7, $s->{WTD_NET_TY},$border1);
			$worksheet->write($a,8, $s->{WTD_NET_LY},$border1);
			$worksheet->write($a,10, $s->{WTD_TARGET},$border1);
				if ($s->{WTD_NET_LY} <= 0){
					$worksheet->write($a,9, "",$subt); }
				else{
					$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$subt); }
										
				if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
					$worksheet->write($a,11, "",$subt); }
				else{
					$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$subt); }
										
			$worksheet->write($a,12, $s->{MTD_NET_TY},$border1);
			$worksheet->write($a,13, $s->{MTD_NET_LY},$border1);
			$worksheet->write($a,15, $s->{MTD_TARGET},$border1);
				if ($s->{MTD_NET_LY} <= 0){
					$worksheet->write($a,14, "",$subt); }
				else{
					$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$subt); }
										
				if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
					$worksheet->write($a,16, "",$subt); }
				else{
					$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$subt); }
											
			$worksheet->write($a,17, $s->{QTD_NET_TY},$border1);
			$worksheet->write($a,18, $s->{QTD_NET_LY},$border1);
			$worksheet->write($a,20, $s->{QTD_TARGET},$border1);
				if ($s->{QTD_NET_LY} <= 0){
					$worksheet->write($a,19, "",$subt); }
				else{
					$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$subt); }
										
				if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
					$worksheet->write($a,21, "",$subt); }
				else{
					$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$subt); }
											
			$worksheet->write($a,22, $s->{YTD_NET_TY},$border1);
			$worksheet->write($a,23, $s->{YTD_NET_LY},$border1);
			$worksheet->write($a,25, $s->{YTD_TARGET},$border1);
				if ($s->{YTD_NET_LY} <= 0){
					$worksheet->write($a,24, "",$subt); }
				else{
					$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$subt); }
										
				if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
					$worksheet->write($a,26, "",$subt); }
				else{
					$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$subt); }
													
			$a++;
			$counter++;
						
		}
		
		$worksheet->write($a,7, $s->{WTD_NET_TY},$bodyNum);
		$worksheet->write($a,8, $s->{WTD_NET_LY},$bodyNum);
		$worksheet->write($a,10, $s->{WTD_TARGET},$bodyNum);
			if ($s->{WTD_NET_LY} <= 0){
				$worksheet->write($a,9, "",$bodyPct); }
			else{
				$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$bodyPct); }
							
			if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
				$worksheet->write($a,11, "",$bodyPct); }
			else{
				$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$bodyPct); }
							
		$worksheet->write($a,12, $s->{MTD_NET_TY},$bodyNum);
		$worksheet->write($a,13, $s->{MTD_NET_LY},$bodyNum);
		$worksheet->write($a,15, $s->{MTD_TARGET},$bodyNum);
			if ($s->{MTD_NET_LY} <= 0){
				$worksheet->write($a,14, "",$bodyPct); }
			else{
				$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$bodyPct); }
								
			if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
				$worksheet->write($a,16, "",$bodyPct); }
			else{
				$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$bodyPct); }
							
		$worksheet->write($a,17, $s->{QTD_NET_TY},$bodyNum);
		$worksheet->write($a,18, $s->{QTD_NET_LY},$bodyNum);
		$worksheet->write($a,20, $s->{QTD_TARGET},$bodyNum);
			if ($s->{QTD_NET_LY} <= 0){
				$worksheet->write($a,19, "",$bodyPct); }
			else{
				$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$bodyPct); }
							
			if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
				$worksheet->write($a,21, "",$bodyPct); }
			else{
				$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$bodyPct); }
								
		$worksheet->write($a,22, $s->{YTD_NET_TY},$bodyNum);
		$worksheet->write($a,23, $s->{YTD_NET_LY},$bodyNum);
		$worksheet->write($a,25, $s->{YTD_TARGET},$bodyNum);
			if ($s->{YTD_NET_LY} <= 0){
				$worksheet->write($a,24, "",$bodyPct); }
			else{
				$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$bodyPct); }
							
			if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
				$worksheet->write($a,26, "",$bodyPct); }
			else{
				$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$bodyPct); }
		
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $flg_desc, $border2 );
		
		$a++;
		$counter = 0;
	}
	
	$worksheet->write($a,7, $s->{WTD_NET_TY},$bodyNum);
	$worksheet->write($a,8, $s->{WTD_NET_LY},$bodyNum);
	$worksheet->write($a,10, $s->{WTD_TARGET},$bodyNum);
		if ($s->{WTD_NET_LY} <= 0){
			$worksheet->write($a,9, "",$bodyPct); }
		else{
			$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$bodyPct); }
						
		if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
			$worksheet->write($a,11, "",$bodyPct); }
		else{
			$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$bodyPct); }
						
	$worksheet->write($a,12, $s->{MTD_NET_TY},$bodyNum);
	$worksheet->write($a,13, $s->{MTD_NET_LY},$bodyNum);
	$worksheet->write($a,15, $s->{MTD_TARGET},$bodyNum);
		if ($s->{MTD_NET_LY} <= 0){
			$worksheet->write($a,14, "",$bodyPct); }
		else{
			$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$bodyPct); }
							
		if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
			$worksheet->write($a,16, "",$bodyPct); }
		else{
			$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$bodyPct); }
						
	$worksheet->write($a,17, $s->{QTD_NET_TY},$bodyNum);
	$worksheet->write($a,18, $s->{QTD_NET_LY},$bodyNum);
	$worksheet->write($a,20, $s->{QTD_TARGET},$bodyNum);
		if ($s->{QTD_NET_LY} <= 0){
			$worksheet->write($a,19, "",$bodyPct); }
		else{
			$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$bodyPct); }
						
		if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
			$worksheet->write($a,21, "",$bodyPct); }
		else{
			$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$bodyPct); }
								
	$worksheet->write($a,22, $s->{YTD_NET_TY},$bodyNum);
	$worksheet->write($a,23, $s->{YTD_NET_LY},$bodyNum);
	$worksheet->write($a,25, $s->{YTD_TARGET},$bodyNum);
		if ($s->{YTD_NET_LY} <= 0){
			$worksheet->write($a,24, "",$bodyPct); }
		else{
			$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$bodyPct); }
						
		if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
			$worksheet->write($a,26, "",$bodyPct); }
		else{
			$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$bodyPct); }
	
	$worksheet->merge_range( $a, 4, $a, 6, 'Total ' . $store_format, $bodyN );
	$worksheet->merge_range( $format_counter, 3, $a, 3, $store_format, $border2 );
	
	$a++;
}

$sls->finish();
$sls1->finish();
$sls2->finish();

}

#sheet 2
sub query_summary_merchandise{

$sls = $dbh->prepare (qq{
	SELECT SUM(WTD_TARGET_DS) WTD_TARGET_DS, SUM(WTD_NET_TY_DS) WTD_NET_TY_DS, SUM(WTD_NET_LY_DS) WTD_NET_LY_DS, SUM(MTD_TARGET_DS) MTD_TARGET_DS, SUM(MTD_NET_TY_DS) MTD_NET_TY_DS, SUM(MTD_NET_LY_DS) MTD_NET_LY_DS, SUM(QTD_TARGET_DS) QTD_TARGET_DS, SUM(QTD_NET_TY_DS) QTD_NET_TY_DS, SUM(QTD_NET_LY_DS) QTD_NET_LY_DS, SUM(YTD_TARGET_DS) YTD_TARGET_DS, SUM(YTD_NET_TY_DS) YTD_NET_TY_DS, SUM(YTD_NET_LY_DS) YTD_NET_LY_DS, SUM(WTD_TARGET_SU) WTD_TARGET_SU, SUM(WTD_NET_TY_SU) WTD_NET_TY_SU, SUM(WTD_NET_LY_SU) WTD_NET_LY_SU, SUM(MTD_TARGET_SU) MTD_TARGET_SU, SUM(MTD_NET_TY_SU) MTD_NET_TY_SU, SUM(MTD_NET_LY_SU) MTD_NET_LY_SU, SUM(QTD_TARGET_SU) QTD_TARGET_SU, SUM(QTD_NET_TY_SU) QTD_NET_TY_SU, SUM(QTD_NET_LY_SU) QTD_NET_LY_SU, SUM(YTD_TARGET_SU) YTD_TARGET_SU, SUM(YTD_NET_TY_SU) YTD_NET_TY_SU, SUM(YTD_NET_LY_SU) YTD_NET_LY_SU
FROM
(SELECT A.FORMAT, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM
	(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A 
	LEFT JOIN
	(	
	SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_DS, 
							0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
							0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
							0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT_WTD
	WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
		AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)DS
	ON DS.FORMAT_DS = A.FORMAT
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_SU, 
							0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
							0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
							0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT_WTD
	WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
		AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)SU
	ON SU.FORMAT_SU= A.FORMAT
UNION ALL
SELECT A.FORMAT, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_DS,
						0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
						0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT' 
		AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))  
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)DS
	ON DS.FORMAT_DS = A.FORMAT
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_SU,
						0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
						0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)SU
	ON SU.FORMAT_SU = A.FORMAT
UNION ALL
SELECT A.FORMAT, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_DS,
						0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT' 
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)DS
	ON DS.FORMAT_DS = A.FORMAT
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_SU,
						0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld'))) 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)SU
	ON SU.FORMAT_SU = A.FORMAT
UNION ALL	
SELECT A.FORMAT, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
						0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
						SUM(TARGET_SALE_VAL) AS YTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT' 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)DS
	ON DS.FORMAT_DS = A.FORMAT
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
						0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
						SUM(TARGET_SALE_VAL) AS YTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT' 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)SU
	ON SU.FORMAT_SU = A.FORMAT)
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
		SELECT FORMAT, SUM(WTD_TARGET_DS) WTD_TARGET_DS, SUM(WTD_NET_TY_DS) WTD_NET_TY_DS, SUM(WTD_NET_LY_DS) WTD_NET_LY_DS, SUM(MTD_TARGET_DS) MTD_TARGET_DS, SUM(MTD_NET_TY_DS) MTD_NET_TY_DS, SUM(MTD_NET_LY_DS) MTD_NET_LY_DS, SUM(QTD_TARGET_DS) QTD_TARGET_DS, SUM(QTD_NET_TY_DS) QTD_NET_TY_DS, SUM(QTD_NET_LY_DS) QTD_NET_LY_DS, SUM(YTD_TARGET_DS) YTD_TARGET_DS, SUM(YTD_NET_TY_DS) YTD_NET_TY_DS, SUM(YTD_NET_LY_DS) YTD_NET_LY_DS, SUM(WTD_TARGET_SU) WTD_TARGET_SU, SUM(WTD_NET_TY_SU) WTD_NET_TY_SU, SUM(WTD_NET_LY_SU) WTD_NET_LY_SU, SUM(MTD_TARGET_SU) MTD_TARGET_SU, SUM(MTD_NET_TY_SU) MTD_NET_TY_SU, SUM(MTD_NET_LY_SU) MTD_NET_LY_SU, SUM(QTD_TARGET_SU) QTD_TARGET_SU, SUM(QTD_NET_TY_SU) QTD_NET_TY_SU, SUM(QTD_NET_LY_SU) QTD_NET_LY_SU, SUM(YTD_TARGET_SU) YTD_TARGET_SU, SUM(YTD_NET_TY_SU) YTD_NET_TY_SU, SUM(YTD_NET_LY_SU) YTD_NET_LY_SU
		FROM
		(SELECT A.FORMAT, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM
		(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A 
		LEFT JOIN
		(	
		SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, 
								SUM(TARGET_SALE_VAL) AS WTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_DS, 
								0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
								0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
								0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
		FROM METRO_IT_SALES_DEPT_WTD
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)DS
		ON DS.FORMAT_DS = A.FORMAT
		LEFT JOIN
		(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, 
								SUM(TARGET_SALE_VAL) AS WTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_SU, 
								0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
								0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
								0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
		FROM METRO_IT_SALES_DEPT_WTD
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)SU
		ON SU.FORMAT_SU= A.FORMAT
	UNION ALL
	SELECT A.FORMAT, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
		(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A 
		LEFT JOIN
		(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, 
							0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_DS,
							0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
							0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
		FROM METRO_IT_SALES_DEPT 
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT' 
			AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))  
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)DS
		ON DS.FORMAT_DS = A.FORMAT
		LEFT JOIN
		(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, 
							0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_SU,
							0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
							0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
		FROM METRO_IT_SALES_DEPT 
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
			AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)SU
		ON SU.FORMAT_SU = A.FORMAT
	UNION ALL
	SELECT A.FORMAT, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
		(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A 
		LEFT JOIN
		(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, 
							0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
							0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS,
							SUM(TARGET_SALE_VAL) AS QTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_DS,
							0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
		FROM METRO_IT_SALES_DEPT 
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT' 
			AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)DS
		ON DS.FORMAT_DS = A.FORMAT
		LEFT JOIN
		(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, 
							0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
							0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU,
							SUM(TARGET_SALE_VAL) AS QTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_SU,
							0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
		FROM METRO_IT_SALES_DEPT 
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
			AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)SU
		ON SU.FORMAT_SU = A.FORMAT
	UNION ALL	
	SELECT A.FORMAT, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
		(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_SALES_DEPT WHERE STORE_FORMAT <> 3)A 
		LEFT JOIN
		(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, 
							0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
							0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
							0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
							SUM(TARGET_SALE_VAL) AS YTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_DS
		FROM METRO_IT_SALES_DEPT 
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT' 
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)DS
		ON DS.FORMAT_DS = A.FORMAT
		LEFT JOIN
		(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, 
							0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
							0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
							0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
							SUM(TARGET_SALE_VAL) AS YTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_SU
		FROM METRO_IT_SALES_DEPT 
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT' 
		GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END)SU
		ON SU.FORMAT_SU = A.FORMAT)
		GROUP BY FORMAT ORDER BY 1
		});
	$sls1->execute();
						
		while(my $s = $sls1->fetchrow_hashref()){
								
		$worksheet->merge_range( $a, 4, $a, 6, $s->{FORMAT}, $desc );
		$worksheet->write($a,7, $s->{WTD_NET_TY_DS},$border1);
		$worksheet->write($a,8, $s->{WTD_NET_LY_DS},$border1);
		$worksheet->write($a,10, $s->{WTD_TARGET_DS},$border1);
			if ($s->{WTD_NET_LY_DS} <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, ($s->{WTD_NET_TY_DS}-$s->{WTD_NET_LY_DS})/$s->{WTD_NET_LY_DS},$subt); }
									
			if ($s->{WTD_NET_TY_DS} <= 0 or $s->{WTD_TARGET_DS} <= 0 ){
				$worksheet->write($a,11, "",$subt); }
			else{
				$worksheet->write($a,11, ($s->{WTD_NET_TY_DS}-$s->{WTD_TARGET_DS})/$s->{WTD_TARGET_DS},$subt); }
				
		$worksheet->write($a,12, $s->{WTD_NET_TY_SU},$border1);
		$worksheet->write($a,13, $s->{WTD_NET_LY_SU},$border1);
		$worksheet->write($a,15, $s->{WTD_TARGET_SU},$border1);
			if ($s->{WTD_NET_LY_SU} <= 0){
				$worksheet->write($a,14, "",$subt); }
			else{
				$worksheet->write($a,14, ($s->{WTD_NET_TY_SU}-$s->{WTD_NET_LY_SU})/$s->{WTD_NET_LY_SU},$subt); }
									
			if ($s->{WTD_NET_TY_SU} <= 0 or $s->{WTD_TARGET_SU} <= 0 ){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, ($s->{WTD_NET_TY_SU}-$s->{WTD_TARGET_SU})/$s->{WTD_TARGET_SU},$subt); }
									
		$worksheet->write($a,17, $s->{MTD_NET_TY_DS},$border1);
		$worksheet->write($a,18, $s->{MTD_NET_LY_DS},$border1);
		$worksheet->write($a,20, $s->{MTD_TARGET_DS},$border1);
			if ($s->{MTD_NET_LY_DS} <= 0){
				$worksheet->write($a,19, "",$subt); }
			else{
				$worksheet->write($a,19, ($s->{MTD_NET_TY_DS}-$s->{MTD_NET_LY_DS})/$s->{MTD_NET_LY_DS},$subt); }
									
			if ($s->{MTD_NET_TY_DS} <= 0 or $s->{MTD_TARGET_DS} <= 0 ){
				$worksheet->write($a,21, "",$subt); }
			else{
				$worksheet->write($a,21, ($s->{MTD_NET_TY_DS}-$s->{MTD_TARGET_DS})/$s->{MTD_TARGET_DS},$subt); }
				
		$worksheet->write($a,22, $s->{MTD_NET_TY_SU},$border1);
		$worksheet->write($a,23, $s->{MTD_NET_LY_SU},$border1);
		$worksheet->write($a,25, $s->{MTD_TARGET_SU},$border1);
			if ($s->{MTD_NET_LY_SU} <= 0){
				$worksheet->write($a,24, "",$subt); }
			else{
				$worksheet->write($a,24, ($s->{MTD_NET_TY_SU}-$s->{MTD_NET_LY_SU})/$s->{MTD_NET_LY_SU},$subt); }
									
			if ($s->{MTD_NET_TY_SU} <= 0 or $s->{MTD_TARGET_SU} <= 0 ){
				$worksheet->write($a,26, "",$subt); }
			else{
				$worksheet->write($a,26, ($s->{MTD_NET_TY_SU}-$s->{MTD_TARGET_SU})/$s->{MTD_TARGET_SU},$subt); }
										
		$worksheet->write($a,27, $s->{QTD_NET_TY_DS},$border1);
		$worksheet->write($a,28, $s->{QTD_NET_LY_DS},$border1);
		$worksheet->write($a,30, $s->{QTD_TARGET_DS},$border1);
			if ($s->{QTD_NET_LY_DS} <= 0){
				$worksheet->write($a,29, "",$subt); }
			else{
				$worksheet->write($a,29, ($s->{QTD_NET_TY_DS}-$s->{QTD_NET_LY_DS})/$s->{QTD_NET_LY_DS},$subt); }
									
			if ($s->{QTD_NET_TY_DS} <= 0 or $s->{QTD_TARGET_DS} <= 0 ){
				$worksheet->write($a,31, "",$subt); }
			else{
				$worksheet->write($a,31, ($s->{QTD_NET_TY_DS}-$s->{QTD_TARGET_DS})/$s->{QTD_TARGET_DS},$subt); }
				
		$worksheet->write($a,32, $s->{QTD_NET_TY_SU},$border1);
		$worksheet->write($a,33, $s->{QTD_NET_LY_SU},$border1);
		$worksheet->write($a,35, $s->{QTD_TARGET_SU},$border1);
			if ($s->{QTD_NET_LY_SU} <= 0){
				$worksheet->write($a,34, "",$subt); }
			else{
				$worksheet->write($a,34, ($s->{QTD_NET_TY_SU}-$s->{QTD_NET_LY_SU})/$s->{QTD_NET_LY_SU},$subt); }
									
			if ($s->{QTD_NET_TY_SU} <= 0 or $s->{QTD_TARGET_SU} <= 0 ){
				$worksheet->write($a,36, "",$subt); }
			else{
				$worksheet->write($a,36, ($s->{QTD_NET_TY_SU}-$s->{QTD_TARGET_SU})/$s->{QTD_TARGET_SU},$subt); }
										
		$worksheet->write($a,37, $s->{YTD_NET_TY_DS},$border1);
		$worksheet->write($a,38, $s->{YTD_NET_LY_DS},$border1);
		$worksheet->write($a,40, $s->{YTD_TARGET_DS},$border1);
			if ($s->{YTD_NET_LY_DS} <= 0){
				$worksheet->write($a,39, "",$subt); }
			else{
				$worksheet->write($a,39, ($s->{YTD_NET_TY_DS}-$s->{YTD_NET_LY_DS})/$s->{YTD_NET_LY_DS},$subt); }
									
			if ($s->{YTD_NET_TY_DS} <= 0 or $s->{YTD_TARGET_DS} <= 0 ){
				$worksheet->write($a,41, "",$subt); }
			else{
				$worksheet->write($a,41, ($s->{YTD_NET_TY_DS}-$s->{YTD_TARGET_DS})/$s->{YTD_TARGET_DS},$subt); }
				
		$worksheet->write($a,42, $s->{YTD_NET_TY_SU},$border1);
		$worksheet->write($a,43, $s->{YTD_NET_LY_SU},$border1);
		$worksheet->write($a,45, $s->{YTD_TARGET_SU},$border1);
			if ($s->{YTD_NET_LY_SU} <= 0){
				$worksheet->write($a,44, "",$subt); }
			else{
				$worksheet->write($a,44, ($s->{YTD_NET_TY_SU}-$s->{YTD_NET_LY_SU})/$s->{YTD_NET_LY_SU},$subt); }
									
			if ($s->{YTD_NET_TY_SU} <= 0 or $s->{YTD_TARGET_SU} <= 0 ){
				$worksheet->write($a,46, "",$subt); }
			else{
				$worksheet->write($a,46, ($s->{YTD_NET_TY_SU}-$s->{YTD_TARGET_SU})/$s->{YTD_TARGET_SU},$subt); }
												
		$a++;
		$counter++;
					
	}
	
	$worksheet->write($a,7, $s->{WTD_NET_TY_DS},$bodyNum);
	$worksheet->write($a,8, $s->{WTD_NET_LY_DS},$bodyNum);
	$worksheet->write($a,10, $s->{WTD_TARGET_DS},$bodyNum);
		if ($s->{WTD_NET_LY_DS} <= 0){
			$worksheet->write($a,9, "",$bodyPct); }
		else{
			$worksheet->write($a,9, ($s->{WTD_NET_TY_DS}-$s->{WTD_NET_LY_DS})/$s->{WTD_NET_LY_DS},$bodyPct); }
						
		if ($s->{WTD_NET_TY_DS} <= 0 or $s->{WTD_TARGET_DS} <= 0 ){
			$worksheet->write($a,11, "",$bodyPct); }
		else{
			$worksheet->write($a,11, ($s->{WTD_NET_TY_DS}-$s->{WTD_TARGET_DS})/$s->{WTD_TARGET_DS},$bodyPct); }
			
	$worksheet->write($a,12, $s->{WTD_NET_TY_SU},$bodyNum);
	$worksheet->write($a,13, $s->{WTD_NET_LY_SU},$bodyNum);
	$worksheet->write($a,15, $s->{WTD_TARGET_SU},$bodyNum);
		if ($s->{WTD_NET_LY_SU} <= 0){
			$worksheet->write($a,14, "",$bodyPct); }
		else{
			$worksheet->write($a,14, ($s->{WTD_NET_TY_SU}-$s->{WTD_NET_LY_SU})/$s->{WTD_NET_LY_SU},$bodyPct); }
						
		if ($s->{WTD_NET_TY_SU} <= 0 or $s->{WTD_TARGET_SU} <= 0 ){
			$worksheet->write($a,16, "",$bodyPct); }
		else{
			$worksheet->write($a,16, ($s->{WTD_NET_TY_SU}-$s->{WTD_TARGET_SU})/$s->{WTD_TARGET_SU},$bodyPct); }
						
	$worksheet->write($a,17, $s->{MTD_NET_TY_DS},$bodyNum);
	$worksheet->write($a,18, $s->{MTD_NET_LY_DS},$bodyNum);
	$worksheet->write($a,20, $s->{MTD_TARGET_DS},$bodyNum);
		if ($s->{MTD_NET_LY_DS} <= 0){
			$worksheet->write($a,19, "",$bodyPct); }
		else{
			$worksheet->write($a,19, ($s->{MTD_NET_TY_DS}-$s->{MTD_NET_LY_DS})/$s->{MTD_NET_LY_DS},$bodyPct); }
							
		if ($s->{MTD_NET_TY_DS} <= 0 or $s->{MTD_TARGET_DS} <= 0 ){
			$worksheet->write($a,21, "",$bodyPct); }
		else{
			$worksheet->write($a,21, ($s->{MTD_NET_TY_DS}-$s->{MTD_TARGET_DS})/$s->{MTD_TARGET_DS},$bodyPct); }
			
	$worksheet->write($a,22, $s->{MTD_NET_TY_SU},$bodyNum);
	$worksheet->write($a,23, $s->{MTD_NET_LY_SU},$bodyNum);
	$worksheet->write($a,25, $s->{MTD_TARGET_SU},$bodyNum);
		if ($s->{MTD_NET_LY_SU} <= 0){
			$worksheet->write($a,24, "",$bodyPct); }
		else{
			$worksheet->write($a,24, ($s->{MTD_NET_TY_SU}-$s->{MTD_NET_LY_SU})/$s->{MTD_NET_LY_SU},$bodyPct); }
							
		if ($s->{MTD_NET_TY_SU} <= 0 or $s->{MTD_TARGET_SU} <= 0 ){
			$worksheet->write($a,26, "",$bodyPct); }
		else{
			$worksheet->write($a,26, ($s->{MTD_NET_TY_SU}-$s->{MTD_TARGET_SU})/$s->{MTD_TARGET_SU},$bodyPct); }
						
	$worksheet->write($a,27, $s->{QTD_NET_TY_DS},$bodyNum);
	$worksheet->write($a,28, $s->{QTD_NET_LY_DS},$bodyNum);
	$worksheet->write($a,30, $s->{QTD_TARGET_DS},$bodyNum);
		if ($s->{QTD_NET_LY_DS} <= 0){
			$worksheet->write($a,29, "",$bodyPct); }
		else{
			$worksheet->write($a,29, ($s->{QTD_NET_TY_DS}-$s->{QTD_NET_LY_DS})/$s->{QTD_NET_LY_DS},$bodyPct); }
						
		if ($s->{QTD_NET_TY_DS} <= 0 or $s->{QTD_TARGET_DS} <= 0 ){
			$worksheet->write($a,31, "",$bodyPct); }
		else{
			$worksheet->write($a,31, ($s->{QTD_NET_TY_DS}-$s->{QTD_TARGET_DS})/$s->{QTD_TARGET_DS},$bodyPct); }
			
	$worksheet->write($a,32, $s->{QTD_NET_TY_SU},$bodyNum);
	$worksheet->write($a,33, $s->{QTD_NET_LY_SU},$bodyNum);
	$worksheet->write($a,35, $s->{QTD_TARGET_SU},$bodyNum);
		if ($s->{QTD_NET_LY_SU} <= 0){
			$worksheet->write($a,34, "",$bodyPct); }
		else{
			$worksheet->write($a,34, ($s->{QTD_NET_TY_SU}-$s->{QTD_NET_LY_SU})/$s->{QTD_NET_LY_SU},$bodyPct); }
						
		if ($s->{QTD_NET_TY_SU} <= 0 or $s->{QTD_TARGET_SU} <= 0 ){
			$worksheet->write($a,36, "",$bodyPct); }
		else{
			$worksheet->write($a,36, ($s->{QTD_NET_TY_SU}-$s->{QTD_TARGET_SU})/$s->{QTD_TARGET_SU},$bodyPct); }
							
	$worksheet->write($a,37, $s->{YTD_NET_TY_DS},$bodyNum);
	$worksheet->write($a,38, $s->{YTD_NET_LY_DS},$bodyNum);
	$worksheet->write($a,40, $s->{YTD_TARGET_DS},$bodyNum);
		if ($s->{YTD_NET_LY_DS} <= 0){
			$worksheet->write($a,39, "",$bodyPct); }
		else{
			$worksheet->write($a,39, ($s->{YTD_NET_TY_DS}-$s->{YTD_NET_LY_DS})/$s->{YTD_NET_LY_DS},$bodyPct); }
						
		if ($s->{YTD_NET_TY_DS} <= 0 or $s->{YTD_TARGET_DS} <= 0 ){
			$worksheet->write($a,41, "",$bodyPct); }
		else{
			$worksheet->write($a,41, ($s->{YTD_NET_TY_DS}-$s->{YTD_TARGET_DS})/$s->{YTD_TARGET_DS},$bodyPct); }
			
	$worksheet->write($a,42, $s->{YTD_NET_TY_SU},$bodyNum);
	$worksheet->write($a,43, $s->{YTD_NET_LY_SU},$bodyNum);
	$worksheet->write($a,45, $s->{YTD_TARGET_SU},$bodyNum);
		if ($s->{YTD_NET_LY_SU} <= 0){
			$worksheet->write($a,44, "",$bodyPct); }
		else{
			$worksheet->write($a,44, ($s->{YTD_NET_TY_SU}-$s->{YTD_NET_LY_SU})/$s->{YTD_NET_LY_SU},$bodyPct); }
						
		if ($s->{YTD_NET_TY_SU} <= 0 or $s->{YTD_TARGET_SU} <= 0 ){
			$worksheet->write($a,46, "",$bodyPct); }
		else{
			$worksheet->write($a,46, ($s->{YTD_NET_TY_SU}-$s->{YTD_TARGET_SU})/$s->{YTD_TARGET_SU},$bodyPct); }
	
	$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
	$worksheet->merge_range( $a-$counter, 3, $a, 3, $summary_label, $border2 );
	
	$counter = 0;
	$a++;
}

$sls->finish();
$sls1->finish();

}

sub query_summary_merchandise_by_store{

$sls = $dbh->prepare (qq{
	SELECT SUM(WTD_TARGET_DS) WTD_TARGET_DS, SUM(WTD_NET_TY_DS) WTD_NET_TY_DS, SUM(WTD_NET_LY_DS) WTD_NET_LY_DS, SUM(MTD_TARGET_DS) MTD_TARGET_DS, SUM(MTD_NET_TY_DS) MTD_NET_TY_DS, SUM(MTD_NET_LY_DS) MTD_NET_LY_DS, SUM(QTD_TARGET_DS) QTD_TARGET_DS, SUM(QTD_NET_TY_DS) QTD_NET_TY_DS, SUM(QTD_NET_LY_DS) QTD_NET_LY_DS, SUM(YTD_TARGET_DS) YTD_TARGET_DS, SUM(YTD_NET_TY_DS) YTD_NET_TY_DS, SUM(YTD_NET_LY_DS) YTD_NET_LY_DS, SUM(WTD_TARGET_SU) WTD_TARGET_SU, SUM(WTD_NET_TY_SU) WTD_NET_TY_SU, SUM(WTD_NET_LY_SU) WTD_NET_LY_SU, SUM(MTD_TARGET_SU) MTD_TARGET_SU, SUM(MTD_NET_TY_SU) MTD_NET_TY_SU, SUM(MTD_NET_LY_SU) MTD_NET_LY_SU, SUM(QTD_TARGET_SU) QTD_TARGET_SU, SUM(QTD_NET_TY_SU) QTD_NET_TY_SU, SUM(QTD_NET_LY_SU) QTD_NET_LY_SU, SUM(YTD_TARGET_SU) YTD_TARGET_SU, SUM(YTD_NET_TY_SU) YTD_NET_TY_SU, SUM(YTD_NET_LY_SU) YTD_NET_LY_SU
FROM
(SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_DS, 
							0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
							0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
							0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT_WTD
	WHERE MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_SU, 
							0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
							0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
							0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT_WTD
	WHERE MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU= A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG	
UNION ALL
SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_DS,
						0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
						0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_SU,
						0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
						0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU = A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG	
UNION ALL	
SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_DS,
						0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld'))) 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_SU,
						0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU = A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG	
UNION ALL		
SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
						0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
						SUM(TARGET_SALE_VAL) AS YTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT' 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
						0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
						SUM(TARGET_SALE_VAL) AS YTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT' 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU = A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG)	
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{		
		SELECT MATURED_FLG, CASE WHEN MATURED_FLG = 1 THEN 'COMP' ELSE 'NEW' END AS FLG_DESC, 
			SUM(WTD_TARGET_DS) WTD_TARGET_DS, SUM(WTD_NET_TY_DS) WTD_NET_TY_DS, SUM(WTD_NET_LY_DS) WTD_NET_LY_DS, SUM(MTD_TARGET_DS) MTD_TARGET_DS, SUM(MTD_NET_TY_DS) MTD_NET_TY_DS, SUM(MTD_NET_LY_DS) MTD_NET_LY_DS, SUM(QTD_TARGET_DS) QTD_TARGET_DS, SUM(QTD_NET_TY_DS) QTD_NET_TY_DS, SUM(QTD_NET_LY_DS) QTD_NET_LY_DS, SUM(YTD_TARGET_DS) YTD_TARGET_DS, SUM(YTD_NET_TY_DS) YTD_NET_TY_DS, SUM(YTD_NET_LY_DS) YTD_NET_LY_DS, SUM(WTD_TARGET_SU) WTD_TARGET_SU, SUM(WTD_NET_TY_SU) WTD_NET_TY_SU, SUM(WTD_NET_LY_SU) WTD_NET_LY_SU, SUM(MTD_TARGET_SU) MTD_TARGET_SU, SUM(MTD_NET_TY_SU) MTD_NET_TY_SU, SUM(MTD_NET_LY_SU) MTD_NET_LY_SU, SUM(QTD_TARGET_SU) QTD_TARGET_SU, SUM(QTD_NET_TY_SU) QTD_NET_TY_SU, SUM(QTD_NET_LY_SU) QTD_NET_LY_SU, SUM(YTD_TARGET_SU) YTD_TARGET_SU, SUM(YTD_NET_TY_SU) YTD_NET_TY_SU, SUM(YTD_NET_LY_SU) YTD_NET_LY_SU
FROM
(SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_DS, 
							0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
							0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
							0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT_WTD
	WHERE MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_SU, 
							0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
							0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
							0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT_WTD
	WHERE MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU= A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG	
UNION ALL
SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_DS,
						0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
						0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_SU,
						0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
						0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU = A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG	
UNION ALL	
SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_DS,
						0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_SU,
						0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU = A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG	
UNION ALL		
SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
						0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
						SUM(TARGET_SALE_VAL) AS YTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT' 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
						0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
						SUM(TARGET_SALE_VAL) AS YTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT' 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU = A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG)	
	WHERE FORMAT = '$store_format'
	GROUP BY MATURED_FLG, CASE WHEN MATURED_FLG = 1 THEN 'COMP' ELSE 'NEW' END
	ORDER BY MATURED_FLG DESC
	});
	$sls1->execute();
		
		$format_counter = $a;
		while(my $s = $sls1->fetchrow_hashref()){
		$flg = $s->{MATURED_FLG};
		$flg_desc = $s->{FLG_DESC};

		$sls2 = $dbh->prepare (qq{			
			SELECT STORE_CODE, STORE_DESCRIPTION,
				SUM(WTD_TARGET_DS) WTD_TARGET_DS, SUM(WTD_NET_TY_DS) WTD_NET_TY_DS, SUM(WTD_NET_LY_DS) WTD_NET_LY_DS, SUM(MTD_TARGET_DS) MTD_TARGET_DS, SUM(MTD_NET_TY_DS) MTD_NET_TY_DS, SUM(MTD_NET_LY_DS) MTD_NET_LY_DS, SUM(QTD_TARGET_DS) QTD_TARGET_DS, SUM(QTD_NET_TY_DS) QTD_NET_TY_DS, SUM(QTD_NET_LY_DS) QTD_NET_LY_DS, SUM(YTD_TARGET_DS) YTD_TARGET_DS, SUM(YTD_NET_TY_DS) YTD_NET_TY_DS, SUM(YTD_NET_LY_DS) YTD_NET_LY_DS, SUM(WTD_TARGET_SU) WTD_TARGET_SU, SUM(WTD_NET_TY_SU) WTD_NET_TY_SU, SUM(WTD_NET_LY_SU) WTD_NET_LY_SU, SUM(MTD_TARGET_SU) MTD_TARGET_SU, SUM(MTD_NET_TY_SU) MTD_NET_TY_SU, SUM(MTD_NET_LY_SU) MTD_NET_LY_SU, SUM(QTD_TARGET_SU) QTD_TARGET_SU, SUM(QTD_NET_TY_SU) QTD_NET_TY_SU, SUM(QTD_NET_LY_SU) QTD_NET_LY_SU, SUM(YTD_TARGET_SU) YTD_TARGET_SU, SUM(YTD_NET_TY_SU) YTD_NET_TY_SU, SUM(YTD_NET_LY_SU) YTD_NET_LY_SU
FROM
(SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format' AND MATURED_FLG = '$flg')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_DS, 
							0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
							0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
							0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT_WTD
	WHERE MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS WTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS WTD_NET_LY_SU, 
							0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
							0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
							0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT_WTD
	WHERE MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU= A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG	
UNION ALL
SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format' AND MATURED_FLG = '$flg')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_DS,
						0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
						0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS MTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS MTD_NET_LY_SU,
						0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
						0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU = A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG	
UNION ALL	
SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format' AND MATURED_FLG = '$flg')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_DS,
						0 AS YTD_TARGET_DS, 0 AS YTD_NET_TY_DS, 0 AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS QTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS QTD_NET_LY_SU,
						0 AS YTD_TARGET_SU, 0 AS YTD_NET_TY_SU, 0 AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT'
		AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU = A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG	
UNION ALL		
SELECT A.FORMAT, A.MATURED_FLG, A.STORE_CODE, A.STORE_DESCRIPTION, WTD_TARGET_DS, WTD_NET_TY_DS, WTD_NET_LY_DS, MTD_TARGET_DS, MTD_NET_TY_DS, MTD_NET_LY_DS, QTD_TARGET_DS, QTD_NET_TY_DS, QTD_NET_LY_DS, YTD_TARGET_DS, YTD_NET_TY_DS, YTD_NET_LY_DS, WTD_TARGET_SU, WTD_NET_TY_SU, WTD_NET_LY_SU, MTD_TARGET_SU, MTD_NET_TY_SU, MTD_NET_LY_SU, QTD_TARGET_SU, QTD_NET_TY_SU, QTD_NET_LY_SU, YTD_TARGET_SU, YTD_NET_TY_SU, YTD_NET_LY_SU FROM	
	(SELECT FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG FROM (SELECT CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG  FROM METRO_IT_SALES_DEPT GROUP BY CASE WHEN STORE_FORMAT = 3 THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG) WHERE FORMAT = '$store_format' AND MATURED_FLG = '$flg')A 
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_DS, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_DS, 0 AS WTD_NET_TY_DS, 0 AS WTD_NET_LY_DS, 
						0 AS MTD_TARGET_DS, 0 AS MTD_NET_TY_DS, 0 AS MTD_NET_LY_DS, 
						0 AS QTD_TARGET_DS, 0 AS QTD_NET_TY_DS, 0 AS QTD_NET_LY_DS, 
						SUM(TARGET_SALE_VAL) AS YTD_TARGET_DS, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_DS, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_DS
	FROM METRO_IT_SALES_DEPT 
	WHERE (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'DS' AND MERCH_GROUP_CODE <> 'Z_OT' 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)DS
	ON DS.FORMAT_DS = A.FORMAT AND DS.STORE_CODE = A.STORE_CODE AND DS.MATURED_FLG = A.MATURED_FLG
	LEFT JOIN
	(SELECT CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END AS FORMAT_SU, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 
						0 AS WTD_TARGET_SU, 0 AS WTD_NET_TY_SU, 0 AS WTD_NET_LY_SU, 
						0 AS MTD_TARGET_SU, 0 AS MTD_NET_TY_SU, 0 AS MTD_NET_LY_SU, 
						0 AS QTD_TARGET_SU, 0 AS QTD_NET_TY_SU, 0 AS QTD_NET_LY_SU, 
						SUM(TARGET_SALE_VAL) AS YTD_TARGET_SU, SUM(NET_SALE_CON_TY) AS YTD_NET_TY_SU, SUM(NET_SALE_CON_LY) AS YTD_NET_LY_SU
	FROM METRO_IT_SALES_DEPT 
	WHERE (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
		AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND MERCH_GROUP_CODE = 'SU' AND MERCH_GROUP_CODE <> 'Z_OT' 
	GROUP BY CASE WHEN STORE_FORMAT = '3' THEN MERCH_GROUP_DESC ELSE UPPER(STORE_FORMAT_DESC) END, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)SU
	ON SU.FORMAT_SU = A.FORMAT AND SU.STORE_CODE = A.STORE_CODE AND SU.MATURED_FLG = A.MATURED_FLG)	
	WHERE FORMAT = '$store_format'
	GROUP BY STORE_CODE, STORE_DESCRIPTION
	ORDER BY 1
	});
		$sls2->execute();
			
			while(my $s = $sls2->fetchrow_hashref()){
									
			$worksheet->write( $a, 5, $s->{STORE_CODE}, $desc );
			$worksheet->write( $a, 6, $s->{STORE_DESCRIPTION}, $desc );
			$worksheet->write($a,7, $s->{WTD_NET_TY_DS},$border1);
			$worksheet->write($a,8, $s->{WTD_NET_LY_DS},$border1);
			$worksheet->write($a,10, $s->{WTD_TARGET_DS},$border1);
				if ($s->{WTD_NET_LY_DS} <= 0){
					$worksheet->write($a,9, "",$subt); }
				else{
					$worksheet->write($a,9, ($s->{WTD_NET_TY_DS}-$s->{WTD_NET_LY_DS})/$s->{WTD_NET_LY_DS},$subt); }
										
				if ($s->{WTD_NET_TY_DS} <= 0 or $s->{WTD_TARGET_DS} <= 0 ){
					$worksheet->write($a,11, "",$subt); }
				else{
					$worksheet->write($a,11, ($s->{WTD_NET_TY_DS}-$s->{WTD_TARGET_DS})/$s->{WTD_TARGET_DS},$subt); }
					
			$worksheet->write($a,12, $s->{WTD_NET_TY_SU},$border1);
			$worksheet->write($a,13, $s->{WTD_NET_LY_SU},$border1);
			$worksheet->write($a,15, $s->{WTD_TARGET_SU},$border1);
				if ($s->{WTD_NET_LY_SU} <= 0){
					$worksheet->write($a,14, "",$subt); }
				else{
					$worksheet->write($a,14, ($s->{WTD_NET_TY_SU}-$s->{WTD_NET_LY_SU})/$s->{WTD_NET_LY_SU},$subt); }
										
				if ($s->{WTD_NET_TY_SU} <= 0 or $s->{WTD_TARGET_SU} <= 0 ){
					$worksheet->write($a,16, "",$subt); }
				else{
					$worksheet->write($a,16, ($s->{WTD_NET_TY_SU}-$s->{WTD_TARGET_SU})/$s->{WTD_TARGET_SU},$subt); }
										
			$worksheet->write($a,17, $s->{MTD_NET_TY_DS},$border1);
			$worksheet->write($a,18, $s->{MTD_NET_LY_DS},$border1);
			$worksheet->write($a,20, $s->{MTD_TARGET_DS},$border1);
				if ($s->{MTD_NET_LY_DS} <= 0){
					$worksheet->write($a,19, "",$subt); }
				else{
					$worksheet->write($a,19, ($s->{MTD_NET_TY_DS}-$s->{MTD_NET_LY_DS})/$s->{MTD_NET_LY_DS},$subt); }
										
				if ($s->{MTD_NET_TY_DS} <= 0 or $s->{MTD_TARGET_DS} <= 0 ){
					$worksheet->write($a,21, "",$subt); }
				else{
					$worksheet->write($a,21, ($s->{MTD_NET_TY_DS}-$s->{MTD_TARGET_DS})/$s->{MTD_TARGET_DS},$subt); }
					
			$worksheet->write($a,22, $s->{MTD_NET_TY_SU},$border1);
			$worksheet->write($a,23, $s->{MTD_NET_LY_SU},$border1);
			$worksheet->write($a,25, $s->{MTD_TARGET_SU},$border1);
				if ($s->{MTD_NET_LY_SU} <= 0){
					$worksheet->write($a,24, "",$subt); }
				else{
					$worksheet->write($a,24, ($s->{MTD_NET_TY_SU}-$s->{MTD_NET_LY_SU})/$s->{MTD_NET_LY_SU},$subt); }
										
				if ($s->{MTD_NET_TY_SU} <= 0 or $s->{MTD_TARGET_SU} <= 0 ){
					$worksheet->write($a,26, "",$subt); }
				else{
					$worksheet->write($a,26, ($s->{MTD_NET_TY_SU}-$s->{MTD_TARGET_SU})/$s->{MTD_TARGET_SU},$subt); }
											
			$worksheet->write($a,27, $s->{QTD_NET_TY_DS},$border1);
			$worksheet->write($a,28, $s->{QTD_NET_LY_DS},$border1);
			$worksheet->write($a,30, $s->{QTD_TARGET_DS},$border1);
				if ($s->{QTD_NET_LY_DS} <= 0){
					$worksheet->write($a,29, "",$subt); }
				else{
					$worksheet->write($a,29, ($s->{QTD_NET_TY_DS}-$s->{QTD_NET_LY_DS})/$s->{QTD_NET_LY_DS},$subt); }
										
				if ($s->{QTD_NET_TY_DS} <= 0 or $s->{QTD_TARGET_DS} <= 0 ){
					$worksheet->write($a,31, "",$subt); }
				else{
					$worksheet->write($a,31, ($s->{QTD_NET_TY_DS}-$s->{QTD_TARGET_DS})/$s->{QTD_TARGET_DS},$subt); }
					
			$worksheet->write($a,32, $s->{QTD_NET_TY_SU},$border1);
			$worksheet->write($a,33, $s->{QTD_NET_LY_SU},$border1);
			$worksheet->write($a,35, $s->{QTD_TARGET_SU},$border1);
				if ($s->{QTD_NET_LY_SU} <= 0){
					$worksheet->write($a,34, "",$subt); }
				else{
					$worksheet->write($a,34, ($s->{QTD_NET_TY_SU}-$s->{QTD_NET_LY_SU})/$s->{QTD_NET_LY_SU},$subt); }
										
				if ($s->{QTD_NET_TY_SU} <= 0 or $s->{QTD_TARGET_SU} <= 0 ){
					$worksheet->write($a,36, "",$subt); }
				else{
					$worksheet->write($a,36, ($s->{QTD_NET_TY_SU}-$s->{QTD_TARGET_SU})/$s->{QTD_TARGET_SU},$subt); }
											
			$worksheet->write($a,37, $s->{YTD_NET_TY_DS},$border1);
			$worksheet->write($a,38, $s->{YTD_NET_LY_DS},$border1);
			$worksheet->write($a,40, $s->{YTD_TARGET_DS},$border1);
				if ($s->{YTD_NET_LY_DS} <= 0){
					$worksheet->write($a,39, "",$subt); }
				else{
					$worksheet->write($a,39, ($s->{YTD_NET_TY_DS}-$s->{YTD_NET_LY_DS})/$s->{YTD_NET_LY_DS},$subt); }
										
				if ($s->{YTD_NET_TY_DS} <= 0 or $s->{YTD_TARGET_DS} <= 0 ){
					$worksheet->write($a,41, "",$subt); }
				else{
					$worksheet->write($a,41, ($s->{YTD_NET_TY_DS}-$s->{YTD_TARGET_DS})/$s->{YTD_TARGET_DS},$subt); }
					
			$worksheet->write($a,42, $s->{YTD_NET_TY_SU},$border1);
			$worksheet->write($a,43, $s->{YTD_NET_LY_SU},$border1);
			$worksheet->write($a,45, $s->{YTD_TARGET_SU},$border1);
				if ($s->{YTD_NET_LY_SU} <= 0){
					$worksheet->write($a,44, "",$subt); }
				else{
					$worksheet->write($a,44, ($s->{YTD_NET_TY_SU}-$s->{YTD_NET_LY_SU})/$s->{YTD_NET_LY_SU},$subt); }
										
				if ($s->{YTD_NET_TY_SU} <= 0 or $s->{YTD_TARGET_SU} <= 0 ){
					$worksheet->write($a,46, "",$subt); }
				else{
					$worksheet->write($a,46, ($s->{YTD_NET_TY_SU}-$s->{YTD_TARGET_SU})/$s->{YTD_TARGET_SU},$subt); }
													
			$a++;
			$counter++;
						
		}
		
		$worksheet->write($a,7, $s->{WTD_NET_TY_DS},$bodyNum);
		$worksheet->write($a,8, $s->{WTD_NET_LY_DS},$bodyNum);
		$worksheet->write($a,10, $s->{WTD_TARGET_DS},$bodyNum);
			if ($s->{WTD_NET_LY_DS} <= 0){
				$worksheet->write($a,9, "",$bodyPct); }
			else{
				$worksheet->write($a,9, ($s->{WTD_NET_TY_DS}-$s->{WTD_NET_LY_DS})/$s->{WTD_NET_LY_DS},$bodyPct); }
							
			if ($s->{WTD_NET_TY_DS} <= 0 or $s->{WTD_TARGET_DS} <= 0 ){
				$worksheet->write($a,11, "",$bodyPct); }
			else{
				$worksheet->write($a,11, ($s->{WTD_NET_TY_DS}-$s->{WTD_TARGET_DS})/$s->{WTD_TARGET_DS},$bodyPct); }
				
		$worksheet->write($a,12, $s->{WTD_NET_TY_SU},$bodyNum);
		$worksheet->write($a,13, $s->{WTD_NET_LY_SU},$bodyNum);
		$worksheet->write($a,15, $s->{WTD_TARGET_SU},$bodyNum);
			if ($s->{WTD_NET_LY_SU} <= 0){
				$worksheet->write($a,14, "",$bodyPct); }
			else{
				$worksheet->write($a,14, ($s->{WTD_NET_TY_SU}-$s->{WTD_NET_LY_SU})/$s->{WTD_NET_LY_SU},$bodyPct); }
							
			if ($s->{WTD_NET_TY_SU} <= 0 or $s->{WTD_TARGET_SU} <= 0 ){
				$worksheet->write($a,16, "",$bodyPct); }
			else{
				$worksheet->write($a,16, ($s->{WTD_NET_TY_SU}-$s->{WTD_TARGET_SU})/$s->{WTD_TARGET_SU},$bodyPct); }
							
		$worksheet->write($a,17, $s->{MTD_NET_TY_DS},$bodyNum);
		$worksheet->write($a,18, $s->{MTD_NET_LY_DS},$bodyNum);
		$worksheet->write($a,20, $s->{MTD_TARGET_DS},$bodyNum);
			if ($s->{MTD_NET_LY_DS} <= 0){
				$worksheet->write($a,19, "",$bodyPct); }
			else{
				$worksheet->write($a,19, ($s->{MTD_NET_TY_DS}-$s->{MTD_NET_LY_DS})/$s->{MTD_NET_LY_DS},$bodyPct); }
								
			if ($s->{MTD_NET_TY_DS} <= 0 or $s->{MTD_TARGET_DS} <= 0 ){
				$worksheet->write($a,21, "",$bodyPct); }
			else{
				$worksheet->write($a,21, ($s->{MTD_NET_TY_DS}-$s->{MTD_TARGET_DS})/$s->{MTD_TARGET_DS},$bodyPct); }
				
		$worksheet->write($a,22, $s->{MTD_NET_TY_SU},$bodyNum);
		$worksheet->write($a,23, $s->{MTD_NET_LY_SU},$bodyNum);
		$worksheet->write($a,25, $s->{MTD_TARGET_SU},$bodyNum);
			if ($s->{MTD_NET_LY_SU} <= 0){
				$worksheet->write($a,24, "",$bodyPct); }
			else{
				$worksheet->write($a,24, ($s->{MTD_NET_TY_SU}-$s->{MTD_NET_LY_SU})/$s->{MTD_NET_LY_SU},$bodyPct); }
								
			if ($s->{MTD_NET_TY_SU} <= 0 or $s->{MTD_TARGET_SU} <= 0 ){
				$worksheet->write($a,26, "",$bodyPct); }
			else{
				$worksheet->write($a,26, ($s->{MTD_NET_TY_SU}-$s->{MTD_TARGET_SU})/$s->{MTD_TARGET_SU},$bodyPct); }
							
		$worksheet->write($a,27, $s->{QTD_NET_TY_DS},$bodyNum);
		$worksheet->write($a,28, $s->{QTD_NET_LY_DS},$bodyNum);
		$worksheet->write($a,30, $s->{QTD_TARGET_DS},$bodyNum);
			if ($s->{QTD_NET_LY_DS} <= 0){
				$worksheet->write($a,29, "",$bodyPct); }
			else{
				$worksheet->write($a,29, ($s->{QTD_NET_TY_DS}-$s->{QTD_NET_LY_DS})/$s->{QTD_NET_LY_DS},$bodyPct); }
							
			if ($s->{QTD_NET_TY_DS} <= 0 or $s->{QTD_TARGET_DS} <= 0 ){
				$worksheet->write($a,31, "",$bodyPct); }
			else{
				$worksheet->write($a,31, ($s->{QTD_NET_TY_DS}-$s->{QTD_TARGET_DS})/$s->{QTD_TARGET_DS},$bodyPct); }
				
		$worksheet->write($a,32, $s->{QTD_NET_TY_SU},$bodyNum);
		$worksheet->write($a,33, $s->{QTD_NET_LY_SU},$bodyNum);
		$worksheet->write($a,35, $s->{QTD_TARGET_SU},$bodyNum);
			if ($s->{QTD_NET_LY_SU} <= 0){
				$worksheet->write($a,34, "",$bodyPct); }
			else{
				$worksheet->write($a,34, ($s->{QTD_NET_TY_SU}-$s->{QTD_NET_LY_SU})/$s->{QTD_NET_LY_SU},$bodyPct); }
							
			if ($s->{QTD_NET_TY_SU} <= 0 or $s->{QTD_TARGET_SU} <= 0 ){
				$worksheet->write($a,36, "",$bodyPct); }
			else{
				$worksheet->write($a,36, ($s->{QTD_NET_TY_SU}-$s->{QTD_TARGET_SU})/$s->{QTD_TARGET_SU},$bodyPct); }
								
		$worksheet->write($a,37, $s->{YTD_NET_TY_DS},$bodyNum);
		$worksheet->write($a,38, $s->{YTD_NET_LY_DS},$bodyNum);
		$worksheet->write($a,40, $s->{YTD_TARGET_DS},$bodyNum);
			if ($s->{YTD_NET_LY_DS} <= 0){
				$worksheet->write($a,39, "",$bodyPct); }
			else{
				$worksheet->write($a,39, ($s->{YTD_NET_TY_DS}-$s->{YTD_NET_LY_DS})/$s->{YTD_NET_LY_DS},$bodyPct); }
							
			if ($s->{YTD_NET_TY_DS} <= 0 or $s->{YTD_TARGET_DS} <= 0 ){
				$worksheet->write($a,41, "",$bodyPct); }
			else{
				$worksheet->write($a,41, ($s->{YTD_NET_TY_DS}-$s->{YTD_TARGET_DS})/$s->{YTD_TARGET_DS},$bodyPct); }
				
		$worksheet->write($a,42, $s->{YTD_NET_TY_SU},$bodyNum);
		$worksheet->write($a,43, $s->{YTD_NET_LY_SU},$bodyNum);
		$worksheet->write($a,45, $s->{YTD_TARGET_SU},$bodyNum);
			if ($s->{YTD_NET_LY_SU} <= 0){
				$worksheet->write($a,44, "",$bodyPct); }
			else{
				$worksheet->write($a,44, ($s->{YTD_NET_TY_SU}-$s->{YTD_NET_LY_SU})/$s->{YTD_NET_LY_SU},$bodyPct); }
							
			if ($s->{YTD_NET_TY_SU} <= 0 or $s->{YTD_TARGET_SU} <= 0 ){
				$worksheet->write($a,46, "",$bodyPct); }
			else{
				$worksheet->write($a,46, ($s->{YTD_NET_TY_SU}-$s->{YTD_TARGET_SU})/$s->{YTD_TARGET_SU},$bodyPct); }
		
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $flg_desc, $border2 );
		
		$a++;
		$counter = 0;
	}
	
	$worksheet->write($a,7, $s->{WTD_NET_TY_DS},$bodyNum);
	$worksheet->write($a,8, $s->{WTD_NET_LY_DS},$bodyNum);
	$worksheet->write($a,10, $s->{WTD_TARGET_DS},$bodyNum);
		if ($s->{WTD_NET_LY_DS} <= 0){
			$worksheet->write($a,9, "",$bodyPct); }
		else{
			$worksheet->write($a,9, ($s->{WTD_NET_TY_DS}-$s->{WTD_NET_LY_DS})/$s->{WTD_NET_LY_DS},$bodyPct); }
						
		if ($s->{WTD_NET_TY_DS} <= 0 or $s->{WTD_TARGET_DS} <= 0 ){
			$worksheet->write($a,11, "",$bodyPct); }
		else{
			$worksheet->write($a,11, ($s->{WTD_NET_TY_DS}-$s->{WTD_TARGET_DS})/$s->{WTD_TARGET_DS},$bodyPct); }
			
	$worksheet->write($a,12, $s->{WTD_NET_TY_SU},$bodyNum);
	$worksheet->write($a,13, $s->{WTD_NET_LY_SU},$bodyNum);
	$worksheet->write($a,15, $s->{WTD_TARGET_SU},$bodyNum);
		if ($s->{WTD_NET_LY_SU} <= 0){
			$worksheet->write($a,14, "",$bodyPct); }
		else{
			$worksheet->write($a,14, ($s->{WTD_NET_TY_SU}-$s->{WTD_NET_LY_SU})/$s->{WTD_NET_LY_SU},$bodyPct); }
						
		if ($s->{WTD_NET_TY_SU} <= 0 or $s->{WTD_TARGET_SU} <= 0 ){
			$worksheet->write($a,16, "",$bodyPct); }
		else{
			$worksheet->write($a,16, ($s->{WTD_NET_TY_SU}-$s->{WTD_TARGET_SU})/$s->{WTD_TARGET_SU},$bodyPct); }
						
	$worksheet->write($a,17, $s->{MTD_NET_TY_DS},$bodyNum);
	$worksheet->write($a,18, $s->{MTD_NET_LY_DS},$bodyNum);
	$worksheet->write($a,20, $s->{MTD_TARGET_DS},$bodyNum);
		if ($s->{MTD_NET_LY_DS} <= 0){
			$worksheet->write($a,19, "",$bodyPct); }
		else{
			$worksheet->write($a,19, ($s->{MTD_NET_TY_DS}-$s->{MTD_NET_LY_DS})/$s->{MTD_NET_LY_DS},$bodyPct); }
							
		if ($s->{MTD_NET_TY_DS} <= 0 or $s->{MTD_TARGET_DS} <= 0 ){
			$worksheet->write($a,21, "",$bodyPct); }
		else{
			$worksheet->write($a,21, ($s->{MTD_NET_TY_DS}-$s->{MTD_TARGET_DS})/$s->{MTD_TARGET_DS},$bodyPct); }
			
	$worksheet->write($a,22, $s->{MTD_NET_TY_SU},$bodyNum);
	$worksheet->write($a,23, $s->{MTD_NET_LY_SU},$bodyNum);
	$worksheet->write($a,25, $s->{MTD_TARGET_SU},$bodyNum);
		if ($s->{MTD_NET_LY_SU} <= 0){
			$worksheet->write($a,24, "",$bodyPct); }
		else{
			$worksheet->write($a,24, ($s->{MTD_NET_TY_SU}-$s->{MTD_NET_LY_SU})/$s->{MTD_NET_LY_SU},$bodyPct); }
							
		if ($s->{MTD_NET_TY_SU} <= 0 or $s->{MTD_TARGET_SU} <= 0 ){
			$worksheet->write($a,26, "",$bodyPct); }
		else{
			$worksheet->write($a,26, ($s->{MTD_NET_TY_SU}-$s->{MTD_TARGET_SU})/$s->{MTD_TARGET_SU},$bodyPct); }
						
	$worksheet->write($a,27, $s->{QTD_NET_TY_DS},$bodyNum);
	$worksheet->write($a,28, $s->{QTD_NET_LY_DS},$bodyNum);
	$worksheet->write($a,30, $s->{QTD_TARGET_DS},$bodyNum);
		if ($s->{QTD_NET_LY_DS} <= 0){
			$worksheet->write($a,29, "",$bodyPct); }
		else{
			$worksheet->write($a,29, ($s->{QTD_NET_TY_DS}-$s->{QTD_NET_LY_DS})/$s->{QTD_NET_LY_DS},$bodyPct); }
						
		if ($s->{QTD_NET_TY_DS} <= 0 or $s->{QTD_TARGET_DS} <= 0 ){
			$worksheet->write($a,31, "",$bodyPct); }
		else{
			$worksheet->write($a,31, ($s->{QTD_NET_TY_DS}-$s->{QTD_TARGET_DS})/$s->{QTD_TARGET_DS},$bodyPct); }
			
	$worksheet->write($a,32, $s->{QTD_NET_TY_SU},$bodyNum);
	$worksheet->write($a,33, $s->{QTD_NET_LY_SU},$bodyNum);
	$worksheet->write($a,35, $s->{QTD_TARGET_SU},$bodyNum);
		if ($s->{QTD_NET_LY_SU} <= 0){
			$worksheet->write($a,34, "",$bodyPct); }
		else{
			$worksheet->write($a,34, ($s->{QTD_NET_TY_SU}-$s->{QTD_NET_LY_SU})/$s->{QTD_NET_LY_SU},$bodyPct); }
						
		if ($s->{QTD_NET_TY_SU} <= 0 or $s->{QTD_TARGET_SU} <= 0 ){
			$worksheet->write($a,36, "",$bodyPct); }
		else{
			$worksheet->write($a,36, ($s->{QTD_NET_TY_SU}-$s->{QTD_TARGET_SU})/$s->{QTD_TARGET_SU},$bodyPct); }
							
	$worksheet->write($a,37, $s->{YTD_NET_TY_DS},$bodyNum);
	$worksheet->write($a,38, $s->{YTD_NET_LY_DS},$bodyNum);
	$worksheet->write($a,40, $s->{YTD_TARGET_DS},$bodyNum);
		if ($s->{YTD_NET_LY_DS} <= 0){
			$worksheet->write($a,39, "",$bodyPct); }
		else{
			$worksheet->write($a,39, ($s->{YTD_NET_TY_DS}-$s->{YTD_NET_LY_DS})/$s->{YTD_NET_LY_DS},$bodyPct); }
						
		if ($s->{YTD_NET_TY_DS} <= 0 or $s->{YTD_TARGET_DS} <= 0 ){
			$worksheet->write($a,41, "",$bodyPct); }
		else{
			$worksheet->write($a,41, ($s->{YTD_NET_TY_DS}-$s->{YTD_TARGET_DS})/$s->{YTD_TARGET_DS},$bodyPct); }
			
	$worksheet->write($a,42, $s->{YTD_NET_TY_SU},$bodyNum);
	$worksheet->write($a,43, $s->{YTD_NET_LY_SU},$bodyNum);
	$worksheet->write($a,45, $s->{YTD_TARGET_SU},$bodyNum);
		if ($s->{YTD_NET_LY_SU} <= 0){
			$worksheet->write($a,44, "",$bodyPct); }
		else{
			$worksheet->write($a,44, ($s->{YTD_NET_TY_SU}-$s->{YTD_NET_LY_SU})/$s->{YTD_NET_LY_SU},$bodyPct); }
						
		if ($s->{YTD_NET_TY_SU} <= 0 or $s->{YTD_TARGET_SU} <= 0 ){
			$worksheet->write($a,46, "",$bodyPct); }
		else{
			$worksheet->write($a,46, ($s->{YTD_NET_TY_SU}-$s->{YTD_TARGET_SU})/$s->{YTD_TARGET_SU},$bodyPct); }
	
	$worksheet->merge_range( $a, 4, $a, 6, 'Total ' . $store_format, $bodyN );
	$worksheet->merge_range( $format_counter, 3, $a, 3, $store_format, $border2 );
	
	$a++;
}

$sls->finish();
$sls1->finish();
$sls2->finish();

}

# sheet 3
sub query_dept {

$sls = $dbh->prepare (qq{
			SELECT 
				SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
				SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
				SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
				SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
			FROM
				(SELECT 
					SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT_WTD
				WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))				
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT 
				WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 					
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
					SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT 
				WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
					AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))				
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
					SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT 
				WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
					AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL))
			});								 
$sls->execute();

	while(my $s = $sls->fetchrow_hashref()){
				
	$sls1 = $dbh->prepare (qq{
				SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
					SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
					SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
					SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
					SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
				FROM
					(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT_WTD
					WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC
					UNION ALL				
					SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT 
					WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
						AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 	
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC
					UNION ALL				
					SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT 
					WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
						AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
						AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC
					UNION ALL				
					SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
						SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT 
					WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
						AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
						AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL)
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC)
				GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC ORDER BY 1
				});								 
	$sls1->execute();

		while(my $s = $sls1->fetchrow_hashref()){
			$merch_group_code = $s->{MERCH_GROUP_CODE};
			$merch_group_desc = $s->{MERCH_GROUP_DESC}; 
			
			$sls2 = $dbh->prepare (qq{
					SELECT GROUP_CODE, GROUP_DESC, 
						SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
						SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
						SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
						SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
					FROM
						(SELECT GROUP_CODE, GROUP_DESC, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT_WTD
						WHERE MERCH_GROUP_CODE = '$merch_group_code' 
							AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
						GROUP BY GROUP_CODE, GROUP_DESC
						UNION ALL				
						SELECT GROUP_CODE, GROUP_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE MERCH_GROUP_CODE = '$merch_group_code' 
							AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 	
						GROUP BY GROUP_CODE, GROUP_DESC
						UNION ALL				
						SELECT GROUP_CODE, GROUP_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
							SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE MERCH_GROUP_CODE = '$merch_group_code' 
							AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
							AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
						GROUP BY GROUP_CODE, GROUP_DESC
						UNION ALL				
						SELECT GROUP_CODE, GROUP_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
							SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE MERCH_GROUP_CODE = '$merch_group_code' 
							AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
							AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL)
						GROUP BY GROUP_CODE, GROUP_DESC)
					GROUP BY GROUP_CODE, GROUP_DESC ORDER BY 1
					});	
			$sls2->execute();
			
			$mgc_counter = $a;
			while(my $s = $sls2->fetchrow_hashref()){
				$group_code = $s->{GROUP_CODE};
				$group_desc = $s->{GROUP_DESC};
						
				$sls3 = $dbh->prepare (qq{
					SELECT DIVISION, DIVISION_DESC, 
						SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
						SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
						SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
						SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
					FROM
						(SELECT DIVISION, DIVISION_DESC, 
								SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
								0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
								0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
								0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
							FROM METRO_IT_SALES_DEPT_WTD
							WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
							GROUP BY DIVISION, DIVISION_DESC
							UNION ALL				
							SELECT DIVISION, DIVISION_DESC, 
								0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
								SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
								0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
								0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
							FROM METRO_IT_SALES_DEPT 
							WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
								AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 	
							GROUP BY DIVISION, DIVISION_DESC
							UNION ALL				
							SELECT DIVISION, DIVISION_DESC, 
								0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
								0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
								SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
								0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
							FROM METRO_IT_SALES_DEPT 
							WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
								AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
								AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
							GROUP BY DIVISION, DIVISION_DESC
							UNION ALL				
							SELECT DIVISION, DIVISION_DESC, 
								0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
								0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
								0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
								SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
							FROM METRO_IT_SALES_DEPT 
							WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
								AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
								AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL)
							GROUP BY DIVISION, DIVISION_DESC)
					GROUP BY DIVISION, DIVISION_DESC ORDER BY 1
					});
				$sls3->execute();
				
				$grp_counter = $a;
				while(my $s = $sls3->fetchrow_hashref()){
					$division = $s->{DIVISION};
					$division_desc = $s->{DIVISION_DESC};
					
					$sls4 = $dbh->prepare (qq{	 
						SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
							SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
							SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
							SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
							SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
						FROM
							(SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
								SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
								0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
								0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
								0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
							FROM METRO_IT_SALES_DEPT_WTD
							WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
							GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC
							UNION ALL				
							SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
								0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
								SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
								0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
								0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
							FROM METRO_IT_SALES_DEPT 
							WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
								AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 	
							GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC
							UNION ALL				
							SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
								0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
								0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
								SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
								0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
							FROM METRO_IT_SALES_DEPT 
							WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
								AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
								AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
							GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC
							UNION ALL				
							SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
								0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
								0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
								0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
								SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
							FROM METRO_IT_SALES_DEPT 
							WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
								AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
								AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL)
							GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC)
						GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC ORDER BY 1
						});
					$sls4->execute();
					
					while(my $s = $sls4->fetchrow_hashref()){
						
						$worksheet->write($a,5, $s->{DEPARTMENT_CODE},$desc);
						$worksheet->write($a,6, $s->{DEPARTMENT_DESC},$desc);
						$worksheet->write($a,7, $s->{WTD_NET_TY},$border1);
						$worksheet->write($a,8, $s->{WTD_NET_LY},$border1);
						$worksheet->write($a,10, $s->{WTD_TARGET},$border1);
							if ($s->{WTD_NET_LY} <= 0){
								$worksheet->write($a,9, "",$subt); }
							else{
								$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$subt); }
						
							if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
								$worksheet->write($a,11, "",$subt); }
							else{
								$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$subt); }
							
						$worksheet->write($a,12, $s->{MTD_NET_TY},$border1);
						$worksheet->write($a,13, $s->{MTD_NET_LY},$border1);
						$worksheet->write($a,15, $s->{MTD_TARGET},$border1);
							if ($s->{MTD_NET_LY} <= 0){
								$worksheet->write($a,14, "",$subt); }
							else{
								$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$subt); }
						
							if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
								$worksheet->write($a,16, "",$subt); }
							else{
								$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$subt); }
								
						$worksheet->write($a,17, $s->{QTD_NET_TY},$border1);
						$worksheet->write($a,18, $s->{QTD_NET_LY},$border1);
						$worksheet->write($a,20, $s->{QTD_TARGET},$border1);
							if ($s->{QTD_NET_LY} <= 0){
								$worksheet->write($a,19, "",$subt); }
							else{
								$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$subt); }
						
							if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
								$worksheet->write($a,21, "",$subt); }
							else{
								$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$subt); }
								
						$worksheet->write($a,22, $s->{YTD_NET_TY},$border1);
						$worksheet->write($a,23, $s->{YTD_NET_LY},$border1);
						$worksheet->write($a,25, $s->{YTD_TARGET},$border1);
							if ($s->{YTD_NET_LY} <= 0){
								$worksheet->write($a,24, "",$subt); }
							else{
								$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$subt); }
						
							if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
								$worksheet->write($a,26, "",$subt); }
							else{
								$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$subt); }
										
						$a++;
						$counter++;
				
					}
					
					$worksheet->write($a,7, $s->{WTD_NET_TY},$bodyNum);
					$worksheet->write($a,8, $s->{WTD_NET_LY},$bodyNum);
					$worksheet->write($a,10, $s->{WTD_TARGET},$bodyNum);
						if ($s->{WTD_NET_LY} <= 0){
							$worksheet->write($a,9, "",$bodyPct); }
						else{
							$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$bodyPct); }
						
						if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
							$worksheet->write($a,11, "",$bodyPct); }
						else{
							$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$bodyPct); }
						
					$worksheet->write($a,12, $s->{MTD_NET_TY},$bodyNum);
					$worksheet->write($a,13, $s->{MTD_NET_LY},$bodyNum);
					$worksheet->write($a,15, $s->{MTD_TARGET},$bodyNum);
						if ($s->{MTD_NET_LY} <= 0){
							$worksheet->write($a,14, "",$bodyPct); }
						else{
							$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$bodyPct); }
							
						if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
							$worksheet->write($a,16, "",$bodyPct); }
						else{
							$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$bodyPct); }
						
					$worksheet->write($a,17, $s->{QTD_NET_TY},$bodyNum);
					$worksheet->write($a,18, $s->{QTD_NET_LY},$bodyNum);
					$worksheet->write($a,20, $s->{QTD_TARGET},$bodyNum);
						if ($s->{QTD_NET_LY} <= 0){
							$worksheet->write($a,19, "",$bodyPct); }
						else{
							$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$bodyPct); }
						
						if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
							$worksheet->write($a,21, "",$bodyPct); }
						else{
							$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$bodyPct); }
							
					$worksheet->write($a,22, $s->{YTD_NET_TY},$bodyNum);
					$worksheet->write($a,23, $s->{YTD_NET_LY},$bodyNum);
					$worksheet->write($a,25, $s->{YTD_TARGET},$bodyNum);
						if ($s->{YTD_NET_LY} <= 0){
							$worksheet->write($a,24, "",$bodyPct); }
						else{
							$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$bodyPct); }
						
						if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
							$worksheet->write($a,26, "",$bodyPct); }
						else{
							$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$bodyPct); }

					$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
					$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
					
					$counter = 0;
					$a++;
				}
				
				$worksheet->write($a,7, $s->{WTD_NET_TY},$bodyNum);
				$worksheet->write($a,8, $s->{WTD_NET_LY},$bodyNum);
				$worksheet->write($a,10, $s->{WTD_TARGET},$bodyNum);
					if ($s->{WTD_NET_LY} <= 0){
						$worksheet->write($a,9, "",$bodyPct); }
					else{
						$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$bodyPct); }

					if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
						$worksheet->write($a,11, "",$bodyPct); }
					else{
						$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$bodyPct); }
					
				$worksheet->write($a,12, $s->{MTD_NET_TY},$bodyNum);
				$worksheet->write($a,13, $s->{MTD_NET_LY},$bodyNum);
				$worksheet->write($a,15, $s->{MTD_TARGET},$bodyNum);
					if ($s->{MTD_NET_LY} <= 0){
						$worksheet->write($a,14, "",$bodyPct); }
					else{
						$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$bodyPct); }
					
					if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
						$worksheet->write($a,16, "",$bodyPct); }
					else{
						$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$bodyPct); }
						
				$worksheet->write($a,17, $s->{QTD_NET_TY},$bodyNum);
				$worksheet->write($a,18, $s->{QTD_NET_LY},$bodyNum);
				$worksheet->write($a,20, $s->{QTD_TARGET},$bodyNum);
					if ($s->{QTD_NET_LY} <= 0){
						$worksheet->write($a,19, "",$bodyPct); }
					else{
						$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$bodyPct); }
					
					if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
						$worksheet->write($a,21, "",$bodyPct); }
					else{
						$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$bodyPct); }
						
				$worksheet->write($a,22, $s->{YTD_NET_TY},$bodyNum);
				$worksheet->write($a,23, $s->{YTD_NET_LY},$bodyNum);
				$worksheet->write($a,25, $s->{YTD_TARGET},$bodyNum);
					if ($s->{YTD_NET_LY} <= 0){
						$worksheet->write($a,24, "",$bodyPct); }
					else{
						$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$bodyPct); }
					
					if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
						$worksheet->write($a,26, "",$bodyPct); }
					else{
						$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$bodyPct); }

				$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
				$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

				$a++;
			}
			
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
			
			$worksheet->write($a,7, $s->{WTD_NET_TY},$headNumber);
			$worksheet->write($a,8, $s->{WTD_NET_LY},$headNumber);
			$worksheet->write($a,10, $s->{WTD_TARGET},$headNumber);
				if ($s->{WTD_NET_LY} <= 0){
					$worksheet->write($a,9, "",$headPct); }
				else{
					$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$headPct); }
					
				if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
					$worksheet->write($a,11, "",$headPct); }
				else{
					$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$headPct); }
			
			$worksheet->write($a,12, $s->{MTD_NET_TY},$headNumber);
			$worksheet->write($a,13, $s->{MTD_NET_LY},$headNumber);
			$worksheet->write($a,15, $s->{MTD_TARGET},$headNumber);
				if ($s->{MTD_NET_LY} <= 0){
					$worksheet->write($a,14, "",$headPct); }
				else{
					$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$headPct); }
					
				if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
					$worksheet->write($a,16, "",$headPct); }
				else{
					$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$headPct); }
					
			$worksheet->write($a,17, $s->{QTD_NET_TY},$headNumber);
			$worksheet->write($a,18, $s->{QTD_NET_LY},$headNumber);
			$worksheet->write($a,20, $s->{QTD_TARGET},$headNumber);
				if ($s->{QTD_NET_LY} <= 0){
					$worksheet->write($a,19, "",$headPct); }
				else{
					$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$headPct); }
					
				if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
					$worksheet->write($a,21, "",$headPct); }
				else{
					$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$headPct); }
					
			$worksheet->write($a,22, $s->{YTD_NET_TY},$headNumber);
			$worksheet->write($a,23, $s->{YTD_NET_LY},$headNumber);
			$worksheet->write($a,25, $s->{YTD_TARGET},$headNumber);
				if ($s->{YTD_NET_LY} <= 0){
					$worksheet->write($a,24, "",$headPct); }
				else{
					$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$headPct); }
					
				if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
					$worksheet->write($a,26, "",$headPct); }
				else{
					$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$headPct); }
			
			$a++;
		}

	$worksheet->write($a,7, $s->{WTD_NET_TY},$headNumber);
	$worksheet->write($a,8, $s->{WTD_NET_LY},$headNumber);
	$worksheet->write($a,10, $s->{WTD_TARGET},$headNumber);
		if ($s->{WTD_NET_LY} <= 0){
			$worksheet->write($a,9, "",$headPct); }
		else{
			$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$headPct); }
			
		if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
			$worksheet->write($a,11, "",$headPct); }
		else{
			$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$headPct); }

	$worksheet->write($a,12, $s->{MTD_NET_TY},$headNumber);
	$worksheet->write($a,13, $s->{MTD_NET_LY},$headNumber);
	$worksheet->write($a,15, $s->{MTD_TARGET},$headNumber);
		if ($s->{MTD_NET_LY} <= 0){
			$worksheet->write($a,14, "",$headPct); }
			else{
			$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$headPct); }
			
		if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$headPct); }
				
	$worksheet->write($a,17, $s->{QTD_NET_TY},$headNumber);
	$worksheet->write($a,18, $s->{QTD_NET_LY},$headNumber);
	$worksheet->write($a,20, $s->{QTD_TARGET},$headNumber);
		if ($s->{QTD_NET_LY} <= 0){
			$worksheet->write($a,19, "",$headPct); }
		else{
			$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$headPct); }
				
		if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
			$worksheet->write($a,21, "",$headPct); }
		else{
			$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$headPct); }
				
	$worksheet->write($a,22, $s->{YTD_NET_TY},$headNumber);
	$worksheet->write($a,23, $s->{YTD_NET_LY},$headNumber);
	$worksheet->write($a,25, $s->{YTD_TARGET},$headNumber);
		if ($s->{YTD_NET_LY} <= 0){
			$worksheet->write($a,24, "",$headPct); }
		else{
			$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$headPct); }
			
		if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
			$worksheet->write($a,26, "",$headPct); }
		else{
			$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$headPct); }
}
	
$worksheet->write($loc, 2, $loc_desc, $bold);
$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls->finish();
$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}

sub query_dept_store {

$sls = $dbh->prepare (qq{
			SELECT 
				SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
				SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
				SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
				SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
			FROM
				(SELECT 
					SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT_WTD
				WHERE STORE_CODE = '$store'
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT 
				WHERE STORE_CODE = '$store'
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
					SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT 
				WHERE STORE_CODE = '$store'
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
					AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
					SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT 
				WHERE STORE_CODE = '$store'
					AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL))
			});								 
$sls->execute();

	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
				SELECT STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
					SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
					SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
					SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
					SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
				FROM
					(SELECT STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT_WTD
					WHERE STORE_CODE = '$store'
					GROUP BY STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC
					UNION ALL				
					SELECT STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT 
					WHERE STORE_CODE = '$store'
						AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
					GROUP BY STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC
					UNION ALL				
					SELECT STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
						SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT 
					WHERE STORE_CODE = '$store'
						AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
						AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
					GROUP BY STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC
					UNION ALL				
					SELECT STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
						SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT 
					WHERE STORE_CODE = '$store'
						AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
						AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL)
					GROUP BY STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC)
				GROUP BY STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC ORDER BY 1, 3
				});								 
	$sls1->execute();

		while(my $s = $sls1->fetchrow_hashref()){
			$merch_group_code = $s->{MERCH_GROUP_CODE};
			$merch_group_desc = $s->{MERCH_GROUP_DESC};
			$loc_code = $s->{STORE_CODE};
			$loc_desc = $s->{STORE_DESCRIPTION};
			
			$sls2 = $dbh->prepare (qq{
					SELECT GROUP_CODE, GROUP_DESC, 
						SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
						SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
						SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
						SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
					FROM
						(SELECT GROUP_CODE, GROUP_DESC, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT_WTD
						WHERE STORE_CODE = '$store' AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE, GROUP_DESC
						UNION ALL				
						SELECT GROUP_CODE, GROUP_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE STORE_CODE = '$store' AND MERCH_GROUP_CODE = '$merch_group_code'
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
						GROUP BY GROUP_CODE, GROUP_DESC
						UNION ALL				
						SELECT GROUP_CODE, GROUP_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
							SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE STORE_CODE = '$store' AND MERCH_GROUP_CODE = '$merch_group_code'
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
							AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
						GROUP BY GROUP_CODE, GROUP_DESC
						UNION ALL				
						SELECT GROUP_CODE, GROUP_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
							SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE STORE_CODE = '$store' AND MERCH_GROUP_CODE = '$merch_group_code'
							AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL)
						GROUP BY GROUP_CODE, GROUP_DESC)
					GROUP BY GROUP_CODE, GROUP_DESC ORDER BY 1
					});	
		$sls2->execute();
		
		$mgc_counter = $a;
		while(my $s = $sls2->fetchrow_hashref()){
			$group_code = $s->{GROUP_CODE};
			$group_desc = $s->{GROUP_DESC};
					
			$sls3 = $dbh->prepare (qq{
				SELECT DIVISION, DIVISION_DESC, 
					SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
					SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
					SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
					SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
				FROM
					(SELECT DIVISION, DIVISION_DESC, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT_WTD
						WHERE STORE_CODE = '$store' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY DIVISION, DIVISION_DESC
						UNION ALL				
						SELECT DIVISION, DIVISION_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE STORE_CODE = '$store' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
						GROUP BY DIVISION, DIVISION_DESC
						UNION ALL				
						SELECT DIVISION, DIVISION_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
							SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE STORE_CODE = '$store' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
							AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
						GROUP BY DIVISION, DIVISION_DESC
						UNION ALL				
						SELECT DIVISION, DIVISION_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
							SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE STORE_CODE = '$store' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL)
						GROUP BY DIVISION, DIVISION_DESC)
				GROUP BY DIVISION, DIVISION_DESC ORDER BY 1
				});
			$sls3->execute();
			
			$grp_counter = $a;
			while(my $s = $sls3->fetchrow_hashref()){
				$division = $s->{DIVISION};
				$division_desc = $s->{DIVISION_DESC};
				
				$sls4 = $dbh->prepare (qq{	 
					SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
						SUM(WTD_TARGET) WTD_TARGET, SUM(WTD_NET_TY) WTD_NET_TY, SUM(WTD_NET_LY) WTD_NET_LY,
						SUM(MTD_TARGET) MTD_TARGET, SUM(MTD_NET_TY) MTD_NET_TY, SUM(MTD_NET_LY) MTD_NET_LY,
						SUM(QTD_TARGET) QTD_TARGET, SUM(QTD_NET_TY) QTD_NET_TY, SUM(QTD_NET_LY) QTD_NET_LY,
						SUM(YTD_TARGET) YTD_TARGET, SUM(YTD_NET_TY) YTD_NET_TY, SUM(YTD_NET_LY) YTD_NET_LY
					FROM
						(SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
							SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_CON_LY) AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT_WTD
						WHERE STORE_CODE = '$store' AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC
						UNION ALL				
						SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_CON_LY) AS MTD_NET_LY,
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE STORE_CODE = '$store' AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
						GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC
						UNION ALL				
						SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
							SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_CON_LY) AS QTD_NET_LY,
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE STORE_CODE = '$store' AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) 
							AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
						GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC
						UNION ALL				
						SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
							SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_CON_LY) AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT 
						WHERE STORE_CODE = '$store' AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE YEAR = '$year' AND DATE_FLD <= TO_DATE('$mo_st_date_fld')))
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL)
						GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC)
					GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC ORDER BY 1
					});
				$sls4->execute();
				
				while(my $s = $sls4->fetchrow_hashref()){
					
					$worksheet->write($a,5, $s->{DEPARTMENT_CODE},$desc);
					$worksheet->write($a,6, $s->{DEPARTMENT_DESC},$desc);
					$worksheet->write($a,7, $s->{WTD_NET_TY},$border1);
					$worksheet->write($a,8, $s->{WTD_NET_LY},$border1);
					$worksheet->write($a,10, $s->{WTD_TARGET},$border1);
					if ($s->{WTD_NET_LY} <= 0){
						$worksheet->write($a,9, "",$subt); }
					else{
						$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$subt); }
					
					if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
						$worksheet->write($a,11, "",$subt); }
					else{
						$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$subt); }
						
					$worksheet->write($a,12, $s->{MTD_NET_TY},$border1);
					$worksheet->write($a,13, $s->{MTD_NET_LY},$border1);
					$worksheet->write($a,15, $s->{MTD_TARGET},$border1);
					if ($s->{MTD_NET_LY} <= 0){
						$worksheet->write($a,14, "",$subt); }
					else{
						$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$subt); }
					
					if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
						$worksheet->write($a,16, "",$subt); }
					else{
						$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$subt); }
						
					$worksheet->write($a,17, $s->{QTD_NET_TY},$border1);
					$worksheet->write($a,18, $s->{QTD_NET_LY},$border1);
					$worksheet->write($a,20, $s->{QTD_TARGET},$border1);
					if ($s->{QTD_NET_LY} <= 0){
						$worksheet->write($a,19, "",$subt); }
					else{
						$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$subt); }
					
					if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
						$worksheet->write($a,21, "",$subt); }
					else{
						$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$subt); }
						
					$worksheet->write($a,22, $s->{YTD_NET_TY},$border1);
					$worksheet->write($a,23, $s->{YTD_NET_LY},$border1);
					$worksheet->write($a,25, $s->{YTD_TARGET},$border1);
					if ($s->{YTD_NET_LY} <= 0){
						$worksheet->write($a,24, "",$subt); }
					else{
						$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$subt); }
					
					if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
						$worksheet->write($a,26, "",$subt); }
					else{
						$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$subt); }
									
					$a++;
					$counter++;
			
				}
				
				$worksheet->write($a,7, $s->{WTD_NET_TY},$bodyNum);
				$worksheet->write($a,8, $s->{WTD_NET_LY},$bodyNum);
				$worksheet->write($a,10, $s->{WTD_TARGET},$bodyNum);
					if ($s->{WTD_NET_LY} <= 0){
						$worksheet->write($a,9, "",$bodyPct); }
					else{
						$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$bodyPct); }
					
					if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
						$worksheet->write($a,11, "",$bodyPct); }
					else{
						$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$bodyPct); }
						
				$worksheet->write($a,12, $s->{MTD_NET_TY},$bodyNum);
				$worksheet->write($a,13, $s->{MTD_NET_LY},$bodyNum);
				$worksheet->write($a,15, $s->{MTD_TARGET},$bodyNum);
					if ($s->{MTD_NET_LY} <= 0){
						$worksheet->write($a,14, "",$bodyPct); }
					else{
						$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$bodyPct); }
					
					if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
						$worksheet->write($a,16, "",$bodyPct); }
					else{
						$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$bodyPct); }
						
				$worksheet->write($a,17, $s->{QTD_NET_TY},$bodyNum);
				$worksheet->write($a,18, $s->{QTD_NET_LY},$bodyNum);
				$worksheet->write($a,20, $s->{QTD_TARGET},$bodyNum);
					if ($s->{QTD_NET_LY} <= 0){
						$worksheet->write($a,19, "",$bodyPct); }
					else{
						$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$bodyPct); }
					
					if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
						$worksheet->write($a,21, "",$bodyPct); }
					else{
						$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$bodyPct); }
						
				$worksheet->write($a,22, $s->{YTD_NET_TY},$bodyNum);
				$worksheet->write($a,23, $s->{YTD_NET_LY},$bodyNum);
				$worksheet->write($a,25, $s->{YTD_TARGET},$bodyNum);
					if ($s->{YTD_NET_LY} <= 0){
						$worksheet->write($a,24, "",$bodyPct); }
					else{
						$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$bodyPct); }
					
					if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
						$worksheet->write($a,26, "",$bodyPct); }
					else{
						$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$bodyPct); }

				$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
				$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
				
				$counter = 0; #RESET dept_counter	
				$a++; #INCREMENT VARIABLE a
			}
			
			$worksheet->write($a,7, $s->{WTD_NET_TY},$bodyNum);
			$worksheet->write($a,8, $s->{WTD_NET_LY},$bodyNum);
			$worksheet->write($a,10, $s->{WTD_TARGET},$bodyNum);
			if ($s->{WTD_NET_LY} <= 0){
				$worksheet->write($a,9, "",$bodyPct); }
			else{
				$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$bodyPct); }

			if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
				$worksheet->write($a,11, "",$bodyPct); }
			else{
				$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$bodyPct); }
				
			$worksheet->write($a,12, $s->{MTD_NET_TY},$bodyNum);
			$worksheet->write($a,13, $s->{MTD_NET_LY},$bodyNum);
			$worksheet->write($a,15, $s->{MTD_TARGET},$bodyNum);
			if ($s->{MTD_NET_LY} <= 0){
				$worksheet->write($a,14, "",$bodyPct); }
			else{
				$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$bodyPct); }
				
			if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
				$worksheet->write($a,16, "",$bodyPct); }
			else{
				$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$bodyPct); }
				
			$worksheet->write($a,17, $s->{QTD_NET_TY},$bodyNum);
			$worksheet->write($a,18, $s->{QTD_NET_LY},$bodyNum);
			$worksheet->write($a,20, $s->{QTD_TARGET},$bodyNum);
			if ($s->{QTD_NET_LY} <= 0){
				$worksheet->write($a,19, "",$bodyPct); }
			else{
				$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$bodyPct); }
				
			if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
				$worksheet->write($a,21, "",$bodyPct); }
			else{
				$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$bodyPct); }
				
			$worksheet->write($a,22, $s->{YTD_NET_TY},$bodyNum);
			$worksheet->write($a,23, $s->{YTD_NET_LY},$bodyNum);
			$worksheet->write($a,25, $s->{YTD_TARGET},$bodyNum);
			if ($s->{YTD_NET_LY} <= 0){
				$worksheet->write($a,24, "",$bodyPct); }
			else{
				$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$bodyPct); }
				
			if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
				$worksheet->write($a,26, "",$bodyPct); }
			else{
				$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{ytd_target})/$s->{YTD_TARGET},$bodyPct); }

			$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

			$a++; #INCREMENT VARIABLE a
		}
		
		if ($merch_group_code eq 'DS'){
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
			$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );}
			
		elsif($merch_group_code eq 'SU'){
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
			$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );}
			
		elsif($merch_group_code eq 'Z_OT'){
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Others', $headN );
			$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'OTHERS', $border2 );}
		
		$worksheet->write($a,7, $s->{WTD_NET_TY},$headNumber);
		$worksheet->write($a,8, $s->{WTD_NET_LY},$headNumber);
		$worksheet->write($a,10, $s->{WTD_TARGET},$headNumber);
			if ($s->{WTD_NET_LY} <= 0){
				$worksheet->write($a,9, "",$headPct); }
			else{
				$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$headPct); }
				
			if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
				$worksheet->write($a,11, "",$headPct); }
			else{
				$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$headPct); }
		
		$worksheet->write($a,12, $s->{MTD_NET_TY},$headNumber);
		$worksheet->write($a,13, $s->{MTD_NET_LY},$headNumber);
		$worksheet->write($a,15, $s->{MTD_TARGET},$headNumber);
			if ($s->{MTD_NET_LY} <= 0){
				$worksheet->write($a,14, "",$headPct); }
			else{
				$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$headPct); }
				
			if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
				$worksheet->write($a,16, "",$headPct); }
			else{
				$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$headPct); }
				
		$worksheet->write($a,17, $s->{QTD_NET_TY},$headNumber);
		$worksheet->write($a,18, $s->{QTD_NET_LY},$headNumber);
		$worksheet->write($a,20, $s->{QTD_TARGET},$headNumber);
			if ($s->{QTD_NET_LY} <= 0){
				$worksheet->write($a,19, "",$headPct); }
			else{
				$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$headPct); }
				
			if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
				$worksheet->write($a,21, "",$headPct); }
			else{
				$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$headPct); }
				
		$worksheet->write($a,22, $s->{YTD_NET_TY},$headNumber);
		$worksheet->write($a,23, $s->{YTD_NET_LY},$headNumber);
		$worksheet->write($a,25, $s->{YTD_TARGET},$headNumber);
			if ($s->{YTD_NET_LY} <= 0){
				$worksheet->write($a,24, "",$headPct); }
			else{
				$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$headPct); }
				
			if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
				$worksheet->write($a,26, "",$headPct); }
			else{
				$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$headPct); }
		
		$a++; #INCREMENT VARIABLE a
	}

	$worksheet->write($a,7, $s->{WTD_NET_TY},$headNumber);
	$worksheet->write($a,8, $s->{WTD_NET_LY},$headNumber);
	$worksheet->write($a,10, $s->{WTD_TARGET},$headNumber);
		if ($s->{WTD_NET_LY} <= 0){
			$worksheet->write($a,9, "",$headPct); }
		else{
			$worksheet->write($a,9, ($s->{WTD_NET_TY}-$s->{WTD_NET_LY})/$s->{WTD_NET_LY},$headPct); }
			
		if ($s->{WTD_NET_TY} <= 0 or $s->{WTD_TARGET} <= 0 ){
			$worksheet->write($a,11, "",$headPct); }
		else{
			$worksheet->write($a,11, ($s->{WTD_NET_TY}-$s->{WTD_TARGET})/$s->{WTD_TARGET},$headPct); }

	$worksheet->write($a,12, $s->{MTD_NET_TY},$headNumber);
	$worksheet->write($a,13, $s->{MTD_NET_LY},$headNumber);
	$worksheet->write($a,15, $s->{MTD_TARGET},$headNumber);
		if ($s->{MTD_NET_LY} <= 0){
			$worksheet->write($a,14, "",$headPct); }
			else{
			$worksheet->write($a,14, ($s->{MTD_NET_TY}-$s->{MTD_NET_LY})/$s->{MTD_NET_LY},$headPct); }
			
		if ($s->{MTD_NET_TY} <= 0 or $s->{MTD_TARGET} <= 0 ){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($s->{MTD_NET_TY}-$s->{MTD_TARGET})/$s->{MTD_TARGET},$headPct); }
				
	$worksheet->write($a,17, $s->{QTD_NET_TY},$headNumber);
	$worksheet->write($a,18, $s->{QTD_NET_LY},$headNumber);
	$worksheet->write($a,20, $s->{QTD_TARGET},$headNumber);
		if ($s->{QTD_NET_LY} <= 0){
			$worksheet->write($a,19, "",$headPct); }
		else{
			$worksheet->write($a,19, ($s->{QTD_NET_TY}-$s->{QTD_NET_LY})/$s->{QTD_NET_LY},$headPct); }
				
		if ($s->{QTD_NET_TY} <= 0 or $s->{QTD_TARGET} <= 0 ){
			$worksheet->write($a,21, "",$headPct); }
		else{
			$worksheet->write($a,21, ($s->{QTD_NET_TY}-$s->{QTD_TARGET})/$s->{QTD_TARGET},$headPct); }
				
	$worksheet->write($a,22, $s->{YTD_NET_TY},$headNumber);
	$worksheet->write($a,23, $s->{YTD_NET_LY},$headNumber);
	$worksheet->write($a,25, $s->{YTD_TARGET},$headNumber);
		if ($s->{YTD_NET_LY} <= 0){
			$worksheet->write($a,24, "",$headPct); }
		else{
			$worksheet->write($a,24, ($s->{YTD_NET_TY}-$s->{YTD_NET_LY})/$s->{YTD_NET_LY},$headPct); }
			
		if ($s->{YTD_NET_TY} <= 0 or $s->{YTD_TARGET} <= 0 ){
			$worksheet->write($a,26, "",$headPct); }
		else{
			$worksheet->write($a,26, ($s->{YTD_NET_TY}-$s->{YTD_TARGET})/$s->{YTD_TARGET},$headPct); }
}

$worksheet->write($loc, 2, $loc_code . " - " . $loc_desc, $bold);			
$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls->finish();
$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}

# SQL to generate raw data
sub generate_csv {

my $truncate = $dbh->prepare( qq{ 
TRUNCATE TABLE METRO_IT_SALES_DEPT_WTD
});
$truncate->execute();

print "Truncated table METRO_IT_SALES_DEPT_WTD... \nPreparing to insert new data... \n";

$test = qq{ 

INSERT INTO METRO_IT_SALES_DEPT_WTD
(STORE_FORMAT, STORE_FORMAT_DESC, STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC, DEPARTMENT_CODE, DEPARTMENT_DESC, 
NEW_FLG, 
MATURED_FLG, 
TARGET_SALE_VAL, 
SALE_TOT_QTY_OTR_TY, NET_SALE_OTR_TY, SALE_TOT_QTY_OTR_LY, NET_SALE_OTR_LY, 
SALE_TOT_QTY_CON_TY, NET_SALE_CON_TY, SALE_TOT_QTY_CON_LY, NET_SALE_CON_LY, UPDATE_DATE)

SELECT 
CASE WHEN STORE_CODE = '2001W' THEN '2' WHEN STORE_CODE = '2005' THEN '4' ELSE STORE_FORMAT END AS STORE_FORMAT, 
CASE WHEN STORE_CODE = '2001W' THEN 'Supermarket' WHEN STORE_CODE = '2005' THEN 'Hypermarket' ELSE STORE_FORMAT_DESC END AS STORE_FORMAT_DESC, 
STORE_CODE, STORE_DESCRIPTION, 
MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
GROUP_CODE, GROUP_DESC, 
DIVISION, DIVISION_DESC, 
DEPARTMENT_CODE, DEPARTMENT_DESC, 
NEW_FLG, MATURED_FLG, 
CASE WHEN SUM(TARGET_SALE_VAL) IS NULL THEN 0 ELSE SUM(TARGET_SALE_VAL) END AS TARGET_SALE_VAL, 
CASE WHEN SUM(SALE_TOT_QTY_OTR_TY) IS NULL THEN 0 ELSE SUM(SALE_TOT_QTY_OTR_TY) END AS SALE_TOT_QTY_OTR_TY, 
CASE WHEN SUM(NET_SALE_OTR_TY) IS NULL THEN 0 ELSE SUM(NET_SALE_OTR_TY) END AS NET_SALE_OTR_TY, 
CASE WHEN SUM(SALE_TOT_QTY_OTR_LY) IS NULL THEN 0 ELSE SUM(SALE_TOT_QTY_OTR_LY) END AS SALE_TOT_QTY_OTR_LY, 
CASE WHEN SUM(NET_SALE_OTR_LY) IS NULL THEN 0 ELSE SUM(NET_SALE_OTR_LY) END AS NET_SALE_OTR_LY, 
CASE WHEN SUM(SALE_TOT_QTY_CON_TY) IS NULL THEN 0 ELSE SUM(SALE_TOT_QTY_CON_TY) END AS SALE_TOT_QTY_CON_TY, 
CASE WHEN SUM(NET_SALE_CON_TY) IS NULL THEN 0 ELSE SUM(NET_SALE_CON_TY) END AS NET_SALE_CON_TY, 
CASE WHEN SUM(SALE_TOT_QTY_CON_LY) IS NULL THEN 0 ELSE SUM(SALE_TOT_QTY_CON_LY) END AS SALE_TOT_QTY_CON_LY, 
CASE WHEN SUM(NET_SALE_CON_LY) IS NULL THEN 0 ELSE SUM(NET_SALE_CON_LY) END AS NET_SALE_CON_LY,
SYSDATE

FROM (
SELECT 
CASE WHEN BASE.STORE_FORMAT = '5' THEN '2' ELSE BASE.STORE_FORMAT END AS STORE_FORMAT, 
CASE WHEN BASE.STORE_FORMAT = '5' THEN 'Supermarket' ELSE BASE.STORE_FORMAT_DESC END AS STORE_FORMAT_DESC, 
BASE.STORE_CODE, 
UPPER(BASE.STORE_DESCRIPTION) AS STORE_DESCRIPTION,
CASE WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 9000) THEN 'Z_OT'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8500) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE = 8040) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE != 8040) THEN 'DS'
ELSE BASE.MERCH_GROUP_CODE END AS MERCH_GROUP_CODE,
CASE WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 9000) THEN 'OTHERS'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8500) THEN 'SUPERMARKET'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE = 8040) THEN 'SUPERMARKET'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE != 8040) THEN 'DEPARTMENT STORE'
ELSE BASE.MERCH_GROUP_DESC END AS MERCH_GROUP_DESC,
BASE.GROUP_CODE,
BASE.GROUP_DESC,
BASE.DIVISION,
BASE.DIVISION_DESC,
BASE.DEPARTMENT_CODE, 
BASE.DEPARTMENT_DESC, 
BASE.NEW_FLG, 
BASE.MATURED_FLG,
SUM(WTD.TARGET_SALE_VAL) TARGET_SALE_VAL,
SUM(NET_SALE_OTR_TY) NET_SALE_OTR_TY, SUM(SALE_TOT_QTY_OTR_TY) SALE_TOT_QTY_OTR_TY, SUM(NET_SALE_CON_TY) NET_SALE_CON_TY, SUM(SALE_TOT_QTY_CON_TY) SALE_TOT_QTY_CON_TY, 
SUM(NET_SALE_OTR_LY) NET_SALE_OTR_LY, SUM(SALE_TOT_QTY_OTR_LY) SALE_TOT_QTY_OTR_LY, SUM(NET_SALE_CON_LY) NET_SALE_CON_LY, SUM(SALE_TOT_QTY_CON_LY) SALE_TOT_QTY_CON_LY
, SUM(NVL(WTD.TARGET_SALE_VAL,0)+NVL(NET_SALE_OTR_TY,0)+NVL(SALE_TOT_QTY_OTR_TY,0)+NVL(NET_SALE_CON_TY,0)+NVL(SALE_TOT_QTY_CON_TY,0)+NVL(NET_SALE_OTR_LY,0)+NVL(SALE_TOT_QTY_OTR_LY,0)+NVL(NET_SALE_CON_LY,0)+NVL(SALE_TOT_QTY_CON_LY,0)) SALE_CHECK
FROM

(SELECT D.MONTH_IN_YEAR, D.PER, S.STORE_KEY, S.STORE_FORMAT, S.STORE_FORMAT_DESC, 
	CASE WHEN S.STORE_CODE IN ('4002') THEN '2001W'
		ELSE S.STORE_CODE END AS STORE_CODE,
		UPPER(S.STORE_DESCRIPTION) AS STORE_DESCRIPTION, 
	M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.GROUP_CODE, M.GROUP_DESC, M.DIVISION, M.DIVISION_DESC, M.DEPARTMENT_CODE, M.DEPARTMENT_DESC,S.NEW_FLG, S.MATURED_FLG
FROM
	(SELECT MONTH_IN_YEAR, TO_CHAR(DATE_FLD, 'MON' ) PER
			FROM DIM_DATE_PRL 
			WHERE DATE_KEY BETWEEN $wk_st_date_key AND $wk_en_date_key 
	GROUP BY MONTH_IN_YEAR, TO_CHAR(DATE_FLD, 'MON' ))D,
	(SELECT DIM.STORE_KEY, DIM.STORE_FORMAT, DIM.STORE_FORMAT_DESC, DIM.STORE_CODE, DIM.STORE_DESCRIPTION, TST.NEW_FLG, TST.MATURED_FLG 
FROM DIM_STORE DIM
  LEFT JOIN (SELECT A.STORE_CODE, A.NEW_FLG, A.MATURED_FLG 
            FROM DIM_STORE A INNER JOIN (SELECT STORE_CODE, MAX(STORE_KEY) STORE_KEY FROM DIM_STORE GROUP BY STORE_CODE)B ON A.STORE_KEY = B.STORE_KEY AND A.STORE_CODE = B.STORE_CODE
            GROUP BY A.STORE_CODE, A.NEW_FLG, A.MATURED_FLG)TST ON DIM.STORE_CODE = TST.STORE_CODE
WHERE DIM.ACTIVE = 1 AND DIM.STORE_FORMAT IN (1, 2, 3, 4, 5) AND DIM.STORE_CODE NOT IN (6008))S,
	(SELECT D.MERCH_GROUP_CODE, D.MERCH_GROUP_DESC, D.GROUP_CODE, D.GROUP_DESC, D.DIVISION, D.DIVISION_DESC, D.DEPARTMENT_CODE, D.DEPARTMENT_DESC 
		FROM DIM_MERCHANDISE M 
			JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE 
								AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
		WHERE D.DIVISION NOT IN (4000))M
GROUP BY D.MONTH_IN_YEAR, D.PER, S.STORE_KEY, S.STORE_FORMAT, S.STORE_FORMAT_DESC, 
	CASE WHEN S.STORE_CODE IN ('4002') THEN '2001W'
		ELSE S.STORE_CODE END,
		UPPER(S.STORE_DESCRIPTION), 
	M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.GROUP_CODE, M.GROUP_DESC, M.DIVISION, M.DIVISION_DESC, M.DEPARTMENT_CODE, M.DEPARTMENT_DESC,S.NEW_FLG, S.MATURED_FLG)BASE
	
LEFT JOIN

(SELECT 
AGG_MLY_STR_DEPT_TARGET.PER, 
DIM_STORE.STORE_KEY, 
DIM_STORE.STORE_FORMAT, 
CASE WHEN DIM_STORE.STORE_CODE IN ('4002') THEN '2001W'
	ELSE DIM_STORE.STORE_CODE END AS STORE_CODE,
DIM_SUB_DEPT.MERCH_GROUP_CODE,
DIM_SUB_DEPT.GROUP_CODE,
DIM_SUB_DEPT.DIVISION,
DIM_SUB_DEPT.DEPARTMENT_CODE,
NVL((SUM(TARGET_SALE_VAL))/1000,0) TARGET_SALE_VAL,  
NVL((SUM((NVL(SALE_NET_VAL_OTR_TY,0))-(NVL(SALE_TOT_TAX_VAL_OTR_TY,0))))/1000,0) NET_SALE_OTR_TY,
SUM(SALE_TOT_QTY_OTR_TY)/1000 SALE_TOT_QTY_OTR_TY, 
NVL((SUM((NVL(SALE_NET_VAL_CON_TY,0))-(NVL(SALE_TOT_TAX_VAL_CON_TY,0))))/1000,0) NET_SALE_CON_TY,
SUM(SALE_TOT_QTY_CON_TY)/1000 SALE_TOT_QTY_CON_TY, 
NVL((SUM((NVL(SALE_NET_VAL_OTR_LY,0))-(NVL(SALE_TOT_TAX_VAL_OTR_LY,0))))/1000,0) NET_SALE_OTR_LY,
SUM(SALE_TOT_QTY_OTR_LY)/1000 SALE_TOT_QTY_OTR_LY, 
NVL((SUM((NVL(SALE_NET_VAL_CON_LY,0))-(NVL(SALE_TOT_TAX_VAL_CON_LY,0))))/1000,0) NET_SALE_CON_LY,
SUM(SALE_TOT_QTY_CON_LY)/1000 SALE_TOT_QTY_CON_LY
FROM (	
	SELECT TBL.PER, TBL.STORE_KEY STORE_KEY, TBL.DS_KEY DS_KEY, TBL.STORE_CODE STORE_CODE, 
		TY_OTR.SALE_NET_VAL_OTR_TY, TY_OTR.SALE_TOT_TAX_VAL_OTR_TY, TY_OTR.SALE_TOT_QTY_OTR_TY,
		TY_CON.SALE_NET_VAL_CON_TY, TY_CON.SALE_TOT_TAX_VAL_CON_TY, TY_CON.SALE_TOT_QTY_CON_TY, 
		LY_OTR.SALE_NET_VAL_OTR_LY, LY_OTR.SALE_TOT_TAX_VAL_OTR_LY, LY_OTR.SALE_TOT_QTY_OTR_LY,
		LY_CON.SALE_NET_VAL_CON_LY, LY_CON.SALE_TOT_TAX_VAL_CON_LY, LY_CON.SALE_TOT_QTY_CON_LY,
		0 AS TARGET_SALE_VAL, 0 AS TARGET_SALE_VAL_LY, 0 AS TARGET_SALE_VAT, 0 AS TARGET_SALE_VAT_LY 
	FROM
		(SELECT D.PER, S.STORE_KEY, S.STORE_CODE, M.DS_KEY
		FROM
			(SELECT TO_CHAR(DATE_FLD, 'MON' ) PER
			FROM DIM_DATE_PRL 
			WHERE DATE_KEY BETWEEN $wk_st_date_key AND $wk_en_date_key 
			GROUP BY TO_CHAR(DATE_FLD, 'MON' ))D,
			(SELECT STORE_KEY, STORE_CODE
			FROM DIM_STORE 
			WHERE ACTIVE = 1 AND STORE_FORMAT IN (1, 2, 3, 4, 5))S,
			(SELECT D.DS_KEY, D.MERCH_GROUP_CODE, D.MERCH_GROUP_DESC, D.GROUP_CODE, D.GROUP_DESC, D.DIVISION, D.DIVISION_DESC, D.DEPARTMENT_CODE, D.DEPARTMENT_DESC
			FROM DIM_MERCHANDISE M 
				JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION 
					AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE )M
		GROUP BY D.PER, S.STORE_KEY, S.STORE_CODE, M.DS_KEY)TBL
		LEFT JOIN
		
		(SELECT TO_CHAR(DA.DATE_FLD, 'MON') PER, 
		STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE, SUM(SALE_TOT_QTY) SALE_TOT_QTY_OTR_TY,
		SUM(SALE_NET_VAL) AS SALE_NET_VAL_OTR_TY, 
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_OTR_TY
		FROM AGG_DLY_STR_PROD AGG 
			INNER JOIN DIM_PRODUCT M ON AGG.PRODUCT_KEY = M.PRODUCT_KEY AND (M.PRODUCT_CATEGORY = 0 OR M.PRODUCT_CATEGORY IS NULL)
			INNER JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
			INNER JOIN DIM_DATE DA ON AGG.DATE_KEY = DA.DATE_KEY
		WHERE AGG.DATE_KEY BETWEEN $wk_st_date_key AND $wk_en_date_key 
		GROUP BY TO_CHAR(DA.DATE_FLD, 'MON'), STORE_KEY, DS_KEY, STORE_CODE)TY_OTR
		
		ON TBL.PER = TY_OTR.PER AND TBL.STORE_KEY = TY_OTR.STORE_KEY AND TBL.STORE_CODE = TY_OTR.STORE_CODE AND TBL.DS_KEY = TY_OTR.DS_KEY
		LEFT JOIN
		
		(SELECT TO_CHAR(DA.DATE_FLD, 'MON') PER, 
		STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE, SUM(SALE_TOT_QTY) SALE_TOT_QTY_CON_TY,  
		SUM(SALE_NET_VAL) AS SALE_NET_VAL_CON_TY, 
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_CON_TY
		FROM AGG_DLY_STR_PROD AGG 
			INNER JOIN DIM_PRODUCT M ON AGG.PRODUCT_KEY = M.PRODUCT_KEY AND M.PRODUCT_CATEGORY = 2
			INNER JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
			INNER JOIN DIM_DATE DA ON AGG.DATE_KEY = DA.DATE_KEY
		WHERE AGG.DATE_KEY BETWEEN $wk_st_date_key AND $wk_en_date_key 
		GROUP BY TO_CHAR(DA.DATE_FLD, 'MON'), STORE_KEY, DS_KEY, STORE_CODE)TY_CON
		
		ON TBL.PER = TY_CON.PER AND TBL.STORE_KEY = TY_CON.STORE_KEY AND TBL.STORE_CODE = TY_CON.STORE_CODE AND TBL.DS_KEY = TY_CON.DS_KEY
		LEFT JOIN
		
		(SELECT TO_CHAR(DA.DATE_FLD, 'MON') PER, 
		STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE, SUM(SALE_TOT_QTY) SALE_TOT_QTY_OTR_LY,  
		SUM(SALE_NET_VAL) AS SALE_NET_VAL_OTR_LY, 
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_OTR_LY
		FROM AGG_DLY_STR_PROD AGG 
			INNER JOIN DIM_PRODUCT M ON AGG.PRODUCT_KEY = M.PRODUCT_KEY AND (M.PRODUCT_CATEGORY = 0 OR M.PRODUCT_CATEGORY IS NULL)
			INNER JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
			INNER JOIN DIM_DATE DA ON AGG.DATE_KEY = DA.DATE_KEY
		WHERE AGG.DATE_KEY BETWEEN $wk_st_date_key_ly AND $wk_en_date_key_ly 
		GROUP BY TO_CHAR(DA.DATE_FLD, 'MON'), STORE_KEY, DS_KEY, STORE_CODE)LY_OTR
		
		ON TBL.PER = LY_OTR.PER AND TBL.STORE_KEY = LY_OTR.STORE_KEY AND TBL.STORE_CODE = LY_OTR.STORE_CODE AND TBL.DS_KEY = LY_OTR.DS_KEY
		LEFT JOIN
		
		(SELECT TO_CHAR(DA.DATE_FLD, 'MON') PER, 
		STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE, SUM(SALE_TOT_QTY) SALE_TOT_QTY_CON_LY,  
		SUM(SALE_NET_VAL) AS SALE_NET_VAL_CON_LY, 
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_CON_LY
		FROM AGG_DLY_STR_PROD AGG 
			INNER JOIN DIM_PRODUCT M ON AGG.PRODUCT_KEY = M.PRODUCT_KEY AND M.PRODUCT_CATEGORY = 2 
			INNER JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
			INNER JOIN DIM_DATE DA ON AGG.DATE_KEY = DA.DATE_KEY
		WHERE AGG.DATE_KEY BETWEEN $wk_st_date_key_ly AND $wk_en_date_key_ly 
		GROUP BY TO_CHAR(DA.DATE_FLD, 'MON'), STORE_KEY, DS_KEY, STORE_CODE)LY_CON
		
		ON TBL.PER = LY_CON.PER AND TBL.STORE_KEY = LY_CON.STORE_KEY AND TBL.STORE_CODE = LY_CON.STORE_CODE AND TBL.DS_KEY = LY_CON.DS_KEY
	
	UNION ALL 
	
	SELECT TO_CHAR(DP.DATE_FLD, 'MON') PER, STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE, 
		0 AS SALE_NET_VAL_OTR_TY, 0 AS SALE_TOT_TAX_VAL_OTR_TY,  0 AS SALE_TOT_QTY_OTR_TY,
		0 AS SALE_NET_VAL_OTR_LY, 0 AS SALE_TOT_TAX_VAL_OTR_LY,  0 AS SALE_TOT_QTY_CON_TY, 
		0 AS SALE_NET_VAL_CON_TY, 0 AS SALE_TOT_TAX_VAL_CON_TY,  0 AS SALE_TOT_QTY_OTR_LY,
		0 AS SALE_NET_VAL_CON_LY, 0 AS SALE_TOT_TAX_VAL_CON_LY,  0 AS SALE_TOT_QTY_CON_LY,		
		SUM (TARGET_SALE_VAL) AS TARGET_SALE_VAL, 
		SUM (TARGET_SALE_VAL_LY) AS TARGET_SALE_VAL_LY, 
		SUM (TARGET_SALE_VAT) AS TARGET_SALE_VAT, 
		SUM (TARGET_SALE_VAT_LY) AS TARGET_SALE_VAT_LY 
	FROM FCT_TARGET A JOIN DIM_DATE_PRL DP ON A.DATE_KEY = DP.DATE_KEY 
	WHERE A.DATE_KEY BETWEEN $wk_st_date_key AND $wk_en_date_key 
	GROUP BY TO_CHAR(DP.DATE_FLD, 'MON'), STORE_KEY, STORE_CODE, DS_KEY 		
		
		) AGG_MLY_STR_DEPT_TARGET,DIM_STORE,DIM_SUB_DEPT
WHERE DIM_STORE.ACTIVE = 1 AND DIM_STORE.STORE_FORMAT IN (1, 2, 3, 4, 5) AND AGG_MLY_STR_DEPT_TARGET.STORE_KEY=DIM_STORE.STORE_KEY AND AGG_MLY_STR_DEPT_TARGET.DS_KEY=DIM_SUB_DEPT.DS_KEY
GROUP BY 
	AGG_MLY_STR_DEPT_TARGET.PER, 
	DIM_STORE.STORE_KEY, DIM_STORE.STORE_FORMAT, 
	CASE WHEN DIM_STORE.STORE_CODE IN ('4002') THEN '2001W'
		ELSE DIM_STORE.STORE_CODE END, 
	DIM_SUB_DEPT.MERCH_GROUP_CODE, DIM_SUB_DEPT.GROUP_CODE, DIM_SUB_DEPT.DIVISION, DIM_SUB_DEPT.DEPARTMENT_CODE
)WTD

ON BASE.PER = WTD.PER AND BASE.STORE_KEY = WTD.STORE_KEY AND BASE.STORE_FORMAT = WTD.STORE_FORMAT AND BASE.STORE_CODE = WTD.STORE_CODE AND BASE.MERCH_GROUP_CODE = WTD.MERCH_GROUP_CODE AND BASE.GROUP_CODE = WTD.GROUP_CODE AND BASE.DIVISION = WTD.DIVISION AND BASE.DEPARTMENT_CODE = WTD.DEPARTMENT_CODE

GROUP BY
CASE WHEN BASE.STORE_FORMAT = '5' THEN '2' ELSE BASE.STORE_FORMAT END, 
CASE WHEN BASE.STORE_FORMAT = '5' THEN 'Supermarket' ELSE BASE.STORE_FORMAT_DESC END, 
BASE.STORE_CODE, 
UPPER(BASE.STORE_DESCRIPTION),
CASE WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 9000) THEN 'Z_OT'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8500) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE = 8040) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE != 8040) THEN 'DS'
ELSE BASE.MERCH_GROUP_CODE END,
CASE WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 9000) THEN 'OTHERS'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8500) THEN 'SUPERMARKET'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE = 8040) THEN 'SUPERMARKET'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE != 8040) THEN 'DEPARTMENT STORE'
ELSE BASE.MERCH_GROUP_DESC END,
BASE.GROUP_CODE, 
BASE.GROUP_DESC, 
BASE.DIVISION, 
BASE.DIVISION_DESC, 
BASE.DEPARTMENT_CODE, 
BASE.DEPARTMENT_DESC, 
BASE.NEW_FLG, 
BASE.MATURED_FLG
)
WHERE SALE_CHECK <> 0
GROUP BY 
CASE WHEN STORE_CODE = '2001W' THEN '2' WHEN STORE_CODE = '2005' THEN '4' ELSE STORE_FORMAT END, 
CASE WHEN STORE_CODE = '2001W' THEN 'Supermarket' WHEN STORE_CODE = '2005' THEN 'Hypermarket' ELSE STORE_FORMAT_DESC END, 
STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC, DEPARTMENT_CODE, DEPARTMENT_DESC, 
NEW_FLG, 
MATURED_FLG
};

my $sth = $dbh->prepare ($test);
$sth->execute;
 
$sth->finish();
$dbh->commit;

print "Done with Insert...";
}

# mailer
sub mail_grp1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1 ) = @ARGV;

#$to = ' chit.lazaro@metrogaisano.com, fili.mercado@metrogaisano.com, karan.malani@metrogaisano.com, julie.montano@metrogaisano.com, lia.chipeco@metrogaisano.com, marlita.portes@metrogaisano.com, jordan.mok@metrogaisano.com, peachy.aquino@metrogaisano.com, patricia.canton@metrogaisano.com, jennifer.yu@metrogaisano.com, jessica.gaisano@metrogaisano.com, april.agapito@metrogaisano.com, edna.prieto@metrogaisano.com, tessie.baldezamo@metrogaisano.com, delia.jakosalem@metrogaisano.com, jennifer.moreno@metrogaisano.com, chedie.lim@metrogaisano.com,glenda.navares@metrogaisano.com, may.sasedor@metrogaisano.com,limuel.ulanday@metrogaisano.com   ';

#$bcc = ' rex.cabanilla@metrogaisano.com, kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, frank.naquines@metrogaisano.com, cham.burgos@metrogaisano.com, roel.gevana@metrogaisano.com, rashel.legaspi@metrogaisano.com';
$bcc = ' lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, frank.naquines@metrogaisano.com, cham.burgos@metrogaisano.com, roel.gevana@metrogaisano.com, rashel.legaspi@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'Daily Sales Performance - Concession as of ' . $as_of;

$msgbody_file = 'message_BI.txt';

$attachment_file_1 = "Daily Sales Performance - Concession (as of $as_of) V1.4.xlsx";


my $msgbody = read_file( $msgbody_file );

my $attachment_data_1 = encode_base64( read_file( $attachment_file_1, 1 ));

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
$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

}

sub mail_grp2 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1 ) = @ARGV;

#$to = ' julie.montano@metrogaisano.com ';

#$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, frank.naquines@metrogaisano.com, cham.burgos@metrogaisano.com ';
		
$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'Daily Sales Performance - Concession as of ' . $as_of;

$msgbody_file = 'message_BI.txt';

$attachment_file_1 = "Daily Sales Performance - Concession (as of $as_of) V1.4.xlsx";

my $msgbody = read_file( $msgbody_file );

my $attachment_data_1 = encode_base64( read_file( $attachment_file_1, 1 ));

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







