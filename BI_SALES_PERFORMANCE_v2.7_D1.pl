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

#change date here  "WHERE DATE_FLD = (SELECT AGG_DLY_END_DATE_FLD FROM ADMIN_ETL_DATE_PARAMETER))

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

	#updated: 2-Mar-16: due to Feb 29 issue, LY has a 1 day spill
	
	$date_3 = qq{ 
	SELECT DATE_KEY1, TO_CHAR(DATE_FLD1+1, 'DD Mon YYYY') DATE_FLD1, DATE_KEY_LY1, TO_CHAR(DATE_FLD_LY1, 'DD Mon YYYY') DATE_FLD_LY1, 
		   DATE_KEY2, TO_CHAR(DATE_FLD2, 'DD Mon YYYY') DATE_FLD2, DATE_KEY_LY2, TO_CHAR(DATE_FLD_LY2, 'DD Mon YYYY') DATE_FLD_LY2,
		   DATE_KEY3, TO_CHAR(DATE_FLD3, 'DD Mon YYYY') DATE_FLD3, DATE_KEY_LY3, TO_CHAR(DATE_FLD_LY3, 'DD Mon YYYY') DATE_FLD_LY3 FROM
		(SELECT DATE_KEY AS DATE_KEY1, DATE_FLD AS DATE_FLD1, DATE_KEY_LY AS DATE_KEY_LY1, DATE_FLD_LY AS DATE_FLD_LY1
		FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_st_date_key-1),
		(SELECT DATE_KEY AS DATE_KEY2, DATE_FLD AS DATE_FLD2, DATE_KEY_LY AS DATE_KEY_LY2, DATE_FLD_LY AS DATE_FLD_LY2
		FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_en_date_key-1),
		(SELECT DATE_KEY AS DATE_KEY3, DATE_FLD AS DATE_FLD3, DATE_KEY_LY AS DATE_KEY_LY3, DATE_FLD_LY AS DATE_FLD_LY3
		FROM DIM_DATE_PRL WHERE QUARTER = $quarter AND YEAR = $year AND MONTH_IN_QUARTER = 1 AND DAY_IN_MONTH = 1)
	 };

	 #$date_3 = qq{ 
	#SELECT DATE_KEY1, TO_CHAR(DATE_FLD1, 'DD Mon YYYY') DATE_FLD1, DATE_KEY_LY1, TO_CHAR(DATE_FLD_LY1, 'DD Mon YYYY') DATE_FLD_LY1, 
		   #DATE_KEY2, TO_CHAR(DATE_FLD2, 'DD Mon YYYY') DATE_FLD2, DATE_KEY_LY2, TO_CHAR(DATE_FLD_LY2, 'DD Mon YYYY') DATE_FLD_LY2,
		   #DATE_KEY3, TO_CHAR(DATE_FLD3, 'DD Mon YYYY') DATE_FLD3, DATE_KEY_LY3, TO_CHAR(DATE_FLD_LY3, 'DD Mon YYYY') DATE_FLD_LY3 FROM
		#(SELECT DATE_KEY AS DATE_KEY1, DATE_FLD AS DATE_FLD1, DATE_KEY_LY AS DATE_KEY_LY1, DATE_FLD_LY AS DATE_FLD_LY1
		#FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_st_date_key),
		#(SELECT DATE_KEY AS DATE_KEY2, DATE_FLD AS DATE_FLD2, DATE_KEY_LY AS DATE_KEY_LY2, DATE_FLD_LY AS DATE_FLD_LY2
		#FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_en_date_key),
		#(SELECT DATE_KEY AS DATE_KEY3, DATE_FLD AS DATE_FLD3, DATE_KEY_LY AS DATE_KEY_LY3, DATE_FLD_LY AS DATE_FLD_LY3
		#FROM DIM_DATE_PRL WHERE QUARTER = $quarter AND YEAR = $year AND MONTH_IN_QUARTER = 1 AND DAY_IN_MONTH = 1)
	 #};

	 
	my $sth_date_3 = $dbh->prepare ($date_3);
	 $sth_date_3->execute;
	 
	while (my $x = $sth_date_3->fetchrow_hashref()) {
		$mo_st_date_key_ly = $x->{DATE_KEY_LY1};
		#$mo_st_date_key_ly = 790;
		$mo_en_date_key_ly = $x->{DATE_KEY_LY2};
		$mo_st_date_fld = $x->{DATE_FLD1};
		$mo_en_date_fld = $x->{DATE_FLD2};
		$mo_st_date_fld_ly = $x->{DATE_FLD_LY1};
		#$mo_st_date_fld_ly = '01 Mar 2015';
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
 
	$workbook = Excel::Writer::XLSX->new("Daily Sales Performance (2101) - Summary (as of $as_of) v2.7.xlsx");
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

	printf "Test ETL Status = ". $test ." \nArc BI Sales Performance (2101) Part 1\nGenerating Data from Source \n";
	
	#&generate_csv;
	&new_sheet_2($sheet = "Department");			
	&call_div;
		
	$workbook->close();
	
	my $pdf_job_1 = Win32::Job->new;
	$pdf_job_1->spawn( "cmd" , q{cmd /C java ecp_FileConverter "Daily Sales Performance (2101) - Summary (as of } . $as_of . q{) v2.7.xlsx" pdf});
	$pdf_job_1->run(60);	
	
	&mail_grp1;	
	
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
#$mo_st_date_fld_head =  $mo_en_date_fld + 1;
#$mo_en_date_fld_head = $x->{DATE_FLD2} + 1;

$worksheet->write($a-11, 2, "Daily Sales Performance (2101)", $bold1);
$worksheet->write($a-10, 2, "WTD: $wk_st_date_key- $wtd_en_dt vs $wtd_st_dt_ly - $wtd_en_dt_ly");
#$worksheet->write($a-10, 2, "WTD: $wtd_st_dt - $wtd_en_dt vs $wtd_st_dt_ly - $wtd_en_dt_ly");
#$worksheet->write($a-9, 2, "MTD: $mo_st_date_fld - $mo_st_date_fld_head vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
#$worksheet->write($a-8, 2, "QTD: $qu_st_date_fld - $mo_st_date_fld_head + 1 vs $qu_st_date_fld_ly - $mo_en_date_fld_ly");
#$worksheet->write($a-7, 2, "YTD: $yr_st_date_fld - $mo_st_date_fld_head + 1 vs $yr_st_date_fld_ly - $mo_en_date_fld_ly");
#Mar-4,2016 : header changed, lacked 1 day
$worksheet->write($a-9, 2, "MTD: $mo_st_date_fld - $wtd_en_dt vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-8, 2, "QTD: $qu_st_date_fld - $wtd_en_dt vs $qu_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-7, 2, "YTD: $yr_st_date_fld - $wtd_en_dt  vs $yr_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-6, 2, "As of $as_of");



##========================= BY STORE ===========================##

foreach my $i ('2101'){ 
# foreach my $i ( '2001', '2001W' ){ 	
	&heading_2;
	&heading;
	&query_dept_store($store = $i);

}

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
$worksheet->set_column( 10, 10, 8 );
$worksheet->set_column( 11, 11, 7 );
# $worksheet->set_column( 10, 11, undef, undef, 1 );
$worksheet->set_column( 12, 13, 10 );
$worksheet->set_column( 14, 14, 7 );
$worksheet->set_column( 15, 15, 10 );
$worksheet->set_column( 16, 16, 7 );
# $worksheet->set_column( 15, 16, undef, undef, 1 );

$worksheet->set_column( 17, 18, 10 );
$worksheet->set_column( 19, 19, 7 );
$worksheet->set_column( 20, 20, 10 );
$worksheet->set_column( 21, 21, 7 );
# $worksheet->set_column( 20, 21, undef, undef, 1 );
$worksheet->set_column( 22, 23, 10 );
$worksheet->set_column( 24, 24, 7 );
$worksheet->set_column( 25, 25, 10 );
$worksheet->set_column( 26, 26, 7 );
# $worksheet->set_column( 25, 26, undef, undef, 1 );

$worksheet->set_column( 27, 28, 10 );
$worksheet->set_column( 29, 29, 7 );
$worksheet->set_column( 30, 30, 10 );
$worksheet->set_column( 31, 31, 7 );
# $worksheet->set_column( 30, 31, undef, undef, 1 );
$worksheet->set_column( 32, 33, 10 );
$worksheet->set_column( 34, 34, 7 );
$worksheet->set_column( 35, 35, 10 );
$worksheet->set_column( 36, 36, 7 );
# $worksheet->set_column( 35, 36, undef, undef, 1 );

$worksheet->set_column( 37, 38, 10 );
$worksheet->set_column( 39, 39, 7 );
$worksheet->set_column( 40, 40, 10 );
$worksheet->set_column( 41, 41, 7 );
# $worksheet->set_column( 40, 41, undef, undef, 1 );
$worksheet->set_column( 42, 43, 10 );
$worksheet->set_column( 44, 44, 7 );
$worksheet->set_column( 45, 45, 10 );
$worksheet->set_column( 46, 46, 7 );
# $worksheet->set_column( 45, 46, undef, undef, 1 );

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
					SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT_WTD
				WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))				
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT 
				WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 					
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
					SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
					SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
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
						SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT_WTD
					WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC
					UNION ALL				
					SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
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
						SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
						SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT 
					WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
						AND (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER < '$quarter' AND YEAR = '$year')
							OR (PER IN (SELECT DISTINCT MONTH_NAME FROM DIM_DATE_PRL WHERE QUARTER = '$quarter' AND YEAR = '$year')))
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
							SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
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
							SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
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
							SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
							SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
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
								SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
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
								SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
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
								SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
								SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
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
								SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
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
								SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
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
								SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
								SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
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
					SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT_WTD
				WHERE STORE_CODE = '$store'
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
					0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY,  
					0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
				FROM METRO_IT_SALES_DEPT 
				WHERE STORE_CODE = '$store'
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) 
				UNION ALL				
				SELECT 
					0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
					0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY,
					SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
					SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
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
						SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
						0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
						0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
						0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
					FROM METRO_IT_SALES_DEPT_WTD
					WHERE STORE_CODE = '$store'
					GROUP BY STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC
					UNION ALL				
					SELECT STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
						0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
						SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
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
						SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
						SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
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
							SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT_WTD
						WHERE STORE_CODE = '$store' AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE, GROUP_DESC
						UNION ALL				
						SELECT GROUP_CODE, GROUP_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
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
							SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
							SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
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
							SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT_WTD
						WHERE STORE_CODE = '$store' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY DIVISION, DIVISION_DESC
						UNION ALL				
						SELECT DIVISION, DIVISION_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
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
							SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
							SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
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
							SUM(TARGET_SALE_VAL) AS WTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS WTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS WTD_NET_LY, 
							0 AS MTD_TARGET, 0 AS MTD_NET_TY, 0 AS MTD_NET_LY, 
							0 AS QTD_TARGET, 0 AS QTD_NET_TY, 0 AS QTD_NET_LY, 
							0 AS YTD_TARGET, 0 AS YTD_NET_TY, 0 AS YTD_NET_LY
						FROM METRO_IT_SALES_DEPT_WTD
						WHERE STORE_CODE = '$store' AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC
						UNION ALL				
						SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
							0 AS WTD_TARGET, 0 AS WTD_NET_TY, 0 AS WTD_NET_LY, 
							SUM(TARGET_SALE_VAL) AS MTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS MTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS MTD_NET_LY,
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
							SUM(TARGET_SALE_VAL) AS QTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS QTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS QTD_NET_LY,
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
							SUM(TARGET_SALE_VAL) AS YTD_TARGET, SUM(NET_SALE_OTR_TY+NET_SALE_CON_TY) AS YTD_NET_TY, SUM(NET_SALE_OTR_LY+NET_SALE_CON_LY) AS YTD_NET_LY
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

# mailer
sub mail_grp1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1, $attachment_file_2 ) = @ARGV;

$to = 'chloy.lamasan@metroretail.com.ph';
$cc = 'lea.gonzaga@metroretail.com.ph,eric.molina@metroretail.com.ph,lloydpatrick.flores@metroretail.com.ph';
#$bcc = 'cham.burgos@metroretail.com.ph,lea.gonzaga@metroretail.com.ph, annalyn.conde@metroretail.com.ph,rex.cabanilla@metroretail.com.ph,eric.molina@metroretail.com.ph';
#$bcc = 'lea.gonzaga@metroretail.com.ph,lloydpatrick.flores@metroretail.com.ph';

$from = 'Report Mailer<report.mailer@metroretail.com.ph>';

$subject = 'Daily Sales Performance (2101) as of ' . $as_of;

$msgbody_file = 'message_BI.txt';

$attachment_file_1 = "Daily Sales Performance (2101) - Summary (as of $as_of) v2.7.xlsx";
$attachment_file_2 = "Daily Sales Performance (2101) - Summary (as of $as_of) v2.7.pdf";

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



