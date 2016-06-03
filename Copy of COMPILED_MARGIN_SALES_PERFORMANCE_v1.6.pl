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
 
	$workbook = Excel::Writer::XLSX->new("TOTAL MARGIN PERFORMANCE (as of $as_of) v1.6.xlsx");
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
	
	#&generate_csv;
	
	&new_sheet($sheet = "Summary");
	&call_str;

	&new_sheet($sheet = "GenMerch_Spmkt");
	&call_str_merchandise;
	
	&new_sheet_2($sheet = "Department");			
	&call_div;
		
	$workbook->close();
	
	my $pdf_job_1 = Win32::Job->new;  
	$pdf_job_1->spawn( "cmd" , q{cmd /C java ecp_FileConverter "TOTAL MARGIN PERFORMANCE (as of } . $as_of . q{) v1.6.xlsx" pdf});
	$pdf_job_1->run(1500);	
	
	&mail_grp1;	
	&mail_external;
	
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

$a = 9, $counter = 0;

$worksheet->write($a-9, 2, "Total Margin Performance", $bold1);
$worksheet->write($a-8, 2, "MTD: $mo_st_date_fld - $mo_en_date_fld");
$worksheet->write($a-7, 2, "As of $as_of");

##========================= COMP STORES ===========================##

&heading_2;
&heading;
&query_dept($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 1, $matured_flg2 = 1, $loc_desc = "COMP STORES");

##========================= ALL STORES ===========================##

$a += 7;
&heading_2;
&heading;
&query_dept($new_flg1 = 0, $new_flg2 = 1, $matured_flg1 = 0, $matured_flg2 = 1, $loc_desc = "ALL STORES");

##========================= BY STORE ===========================##

foreach my $i ( '2001', '2001W', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', '2013', '2223', '3001', '3002', '3003', '3004', '3005', '3006', '3007', '3009', '3010', '3012', '4003', '4004', '6001', '6002', '6003', '6004', '6005', '6009', '6010', '6012', '80001', '80001', '80011', '80031', '80041', '80051', '80061', '80071', '80081', '80101' ){ 
# foreach my $i ( '2001', '2001W' ){ 
	$a += 7;	
	&heading_2;
	&heading;
	&query_dept_store($store = $i);

}

}

sub call_str {

$a=9, $counter=0;

$worksheet->write($a-9, 3, "Total Margin Performance", $bold1);
$worksheet->write($a-8, 3, "MTD: $mo_st_date_fld - $mo_en_date_fld");
$worksheet->write($a-7, 3, "As of $as_of");

$worksheet->write($a-4, 3, "Summary", $bold);

&heading;

$worksheet->merge_range( $a-2, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a-1, 6, 'Format', $subhead );

&query_summary($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 1, $matured_flg2 = 1, $summary_label = 'COMP');
&query_summary($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 0, $matured_flg2 = 0, $summary_label = 'NEW');
&query_summary($new_flg1 = 0, $new_flg2 = 1, $matured_flg1 = 0, $matured_flg2 = 1, $summary_label = 'ALL');

$a+=6; 

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

$a=10, $counter=0;

$worksheet->write($a-10, 3, "Total Margin Performance", $bold1);
$worksheet->write($a-9, 3, "MTD: $mo_st_date_fld - $mo_en_date_fld");
$worksheet->write($a-8, 3, "As of $as_of");

$worksheet->write($a-5, 3, "Summary", $bold);

&heading_3;

$worksheet->merge_range( $a-3, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-3, 4, $a-1, 6, 'Format', $subhead );

&query_summary_merchandise($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 1, $matured_flg2 = 1, $summary_label = 'COMP');
&query_summary_merchandise($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 0, $matured_flg2 = 0, $summary_label = 'NEW');
&query_summary_merchandise($new_flg1 = 0, $new_flg2 = 1, $matured_flg1 = 0, $matured_flg2 = 1, $summary_label = 'ALL');

$a+=7; 

$worksheet->write($a-5, 3, "Per Store", $bold);

&heading_3;

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
#$worksheet->conditional_formatting( 'F9:AK6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );

}

sub new_sheet_2{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom( 92 );
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
#$worksheet->set_print_scale( 100 );
$worksheet->fit_to_pages( 1, 1 );
$worksheet->set_margins( 0.001 );
#$worksheet->conditional_formatting( 'F9:AK6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );
# $worksheet->set_column( 0, 0, undef, undef, 1 );
# $worksheet->set_column( 1, 2, 3 );
# $worksheet->set_column( 3, 4, 4 );
# $worksheet->set_column( 5, 5, undef, undef, 1 );
# $worksheet->set_column( 6, 6, 23 );

# $worksheet->set_column( 7, 8, 8 );
# $worksheet->set_column( 9, 9, 7 );
# $worksheet->set_column( 10, 13, undef, undef, 1 );

# $worksheet->set_column( 14, 15, 8 );
# $worksheet->set_column( 16, 16, 7 );
# $worksheet->set_column( 17, 20, undef, undef, 1 );

# $worksheet->set_column( 21, 22, 8 );
# $worksheet->set_column( 23, 23, 7 );
# $worksheet->set_column( 24, 27, undef, undef, 1 );

}

# headers
sub heading {

$worksheet->write($a-3, 3, "in 000's", $script);
$worksheet->merge_range( $a-4, 7, $a-3, 13, 'TOTAL', $subhead );
$worksheet->merge_range( $a-4, 14, $a-4, 66, 'OUTRIGHT', $subhead );
$worksheet->merge_range( $a-4, 67, $a-4, 98, 'CONCESSION', $subhead );

$worksheet->merge_range( $a-3, 14, $a-3, 20, 'TOTAL - OUTRIGHT', $subhead );
$worksheet->merge_range( $a-3, 21, $a-3, 27, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 28, $a-3, 54, 'FRONT - OTHERS', $subhead );
$worksheet->merge_range( $a-3, 55, $a-3, 63, 'BACK', $subhead );
$worksheet->merge_range( $a-3, 64, $a-3, 66, 'OTHER COST', $subhead );
$worksheet->merge_range( $a-3, 67, $a-3, 73, 'TOTAL - CONCESSION', $subhead );
$worksheet->merge_range( $a-3, 74, $a-3, 80, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 81, $a-3, 83, 'FRONT - OTHERS', $subhead );
$worksheet->merge_range( $a-3, 84, $a-3, 98, 'BACK', $subhead );

foreach my $i ( 7, 14, 21, 67, 74 ) {
	$worksheet->merge_range( $a-2, $i, $a-2, $i+2, 'Actual', $subhead );
	$worksheet->merge_range( $a-2, $i+3, $a-2, $i+4, 'Budget', $subhead );
	$worksheet->merge_range( $a-2, $i+5, $a-2, $i+6, 'Var', $subhead );
	$worksheet->write($a-1, $i, "Sales", $subhead);
	$worksheet->write($a-1, $i+1, "GM", $subhead);
	$worksheet->write($a-1, $i+2, "GM%", $subhead);
	$worksheet->write($a-1, $i+3, "GM", $subhead);
	$worksheet->write($a-1, $i+4, "GM%", $subhead);
	$worksheet->write($a-1, $i+5, "GM", $subhead);
	$worksheet->write($a-1, $i+6, "GM%", $subhead);
	
	$worksheet->set_column( $i, $i+1, 8 );
	$worksheet->set_column( $i+2, $i+2, 7 );
	$worksheet->set_column( $i+3, $i+6, undef, undef, 1 );
}

foreach my $i ( 28, 31, 34, 37, 40, 43, 46, 49, 52, 55, 58, 61, 64, 81, 84, 87, 90, 93, 96 ) {
	$worksheet->write($a-1, $i, 'Actual', $subhead);
	$worksheet->write($a-1, $i+1, 'Budget', $subhead);
	$worksheet->write($a-1, $i+2, 'Var', $subhead);
	
	$worksheet->set_column( $i, $i, 8 );
	$worksheet->set_column( $i+1, $i+2, undef, undef, 1 );
}

$worksheet->merge_range( $a-2, 28, $a-2, 30, 'Cost of Sales - Trade', $subhead );
$worksheet->merge_range( $a-2, 31, $a-2, 33, 'Transfer Discrepancy', $subhead );
$worksheet->merge_range( $a-2, 34, $a-2, 36, 'Promotional Item Charged to Margin', $subhead );
$worksheet->merge_range( $a-2, 37, $a-2, 39, 'Wastage', $subhead );
$worksheet->merge_range( $a-2, 40, $a-2, 42, 'Invoice Price Variance - Trade', $subhead );
$worksheet->merge_range( $a-2, 43, $a-2, 45, 'Shrinkage Cost', $subhead );
$worksheet->merge_range( $a-2, 46, $a-2, 48, 'Cost Variance', $subhead );
$worksheet->merge_range( $a-2, 49, $a-2, 51, 'Synchronization Account', $subhead );
$worksheet->merge_range( $a-2, 52, $a-2, 54, 'Freight Recovery', $subhead );
$worksheet->merge_range( $a-2, 55, $a-2, 57, 'Purchase Allowance', $subhead );
$worksheet->merge_range( $a-2, 58, $a-2, 60, 'Purchase Discouns-Special', $subhead );
$worksheet->merge_range( $a-2, 61, $a-2, 63, 'Other Income-Metro Vendor Portal', $subhead );
$worksheet->merge_range( $a-2, 64, $a-2, 66, 'Freight', $subhead );
$worksheet->merge_range( $a-2, 81, $a-2, 83, 'Cost of Sales - Conc', $subhead );
$worksheet->merge_range( $a-2, 84, $a-2, 86, 'Ad Support', $subhead );
$worksheet->merge_range( $a-2, 87, $a-2, 89, 'Other Income-Storage/Concession', $subhead );
$worksheet->merge_range( $a-2, 90, $a-2, 92, 'Light Recovery', $subhead );
$worksheet->merge_range( $a-2, 93, $a-2, 95, 'Water Recovery', $subhead );
$worksheet->merge_range( $a-2, 96, $a-2, 98, 'Supplies Recovery', $subhead );

$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 23 );
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

$worksheet->write($a-4, 3, "in 000's", $script);

$worksheet->merge_range( $a-5, 7, $a-5, 98, 'GENERAL MERCHANDISE', $subhead );
$worksheet->merge_range( $a-5, 99, $a-5, 190, 'SUPERMARKET', $subhead );

$worksheet->merge_range( $a-4, 7, $a-3, 13, 'TOTAL', $subhead );
$worksheet->merge_range( $a-4, 14, $a-4, 66, 'OUTRIGHT', $subhead );
$worksheet->merge_range( $a-4, 67, $a-4, 98, 'CONCESSION', $subhead );

$worksheet->merge_range( $a-4, 99, $a-3, 105, 'TOTAL', $subhead );
$worksheet->merge_range( $a-4, 106, $a-4, 158, 'OUTRIGHT', $subhead );
$worksheet->merge_range( $a-4, 159, $a-4, 190, 'CONCESSION', $subhead );

$worksheet->merge_range( $a-3, 14, $a-3, 20, 'TOTAL - OUTRIGHT', $subhead );
$worksheet->merge_range( $a-3, 21, $a-3, 27, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 28, $a-3, 54, 'FRONT - OTHERS', $subhead );
$worksheet->merge_range( $a-3, 55, $a-3, 63, 'BACK', $subhead );
$worksheet->merge_range( $a-3, 64, $a-3, 66, 'OTHER COST', $subhead );
$worksheet->merge_range( $a-3, 67, $a-3, 73, 'TOTAL - CONCESSION', $subhead );
$worksheet->merge_range( $a-3, 74, $a-3, 80, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 81, $a-3, 83, 'FRONT - OTHERS', $subhead );
$worksheet->merge_range( $a-3, 84, $a-3, 98, 'BACK', $subhead );

$worksheet->merge_range( $a-3, 106, $a-3, 112, 'TOTAL - OUTRIGHT', $subhead );
$worksheet->merge_range( $a-3, 113, $a-3, 119, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 120, $a-3, 146, 'FRONT - OTHERS', $subhead );
$worksheet->merge_range( $a-3, 147, $a-3, 155, 'BACK', $subhead );
$worksheet->merge_range( $a-3, 156, $a-3, 158, 'OTHER COST', $subhead );
$worksheet->merge_range( $a-3, 159, $a-3, 165, 'TOTAL - CONCESSION', $subhead );
$worksheet->merge_range( $a-3, 166, $a-3, 172, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 173, $a-3, 175, 'FRONT - OTHERS', $subhead );
$worksheet->merge_range( $a-3, 176, $a-3, 190, 'BACK', $subhead );

foreach my $i ( 7, 14, 21, 67, 74, 99, 106, 113, 159, 166 ) {
	$worksheet->merge_range( $a-2, $i, $a-2, $i+2, 'Actual', $subhead );
	$worksheet->merge_range( $a-2, $i+3, $a-2, $i+4, 'Budget', $subhead );
	$worksheet->merge_range( $a-2, $i+5, $a-2, $i+6, 'Var', $subhead );
	$worksheet->write($a-1, $i, "Sales", $subhead);
	$worksheet->write($a-1, $i+1, "GM", $subhead);
	$worksheet->write($a-1, $i+2, "GM%", $subhead);
	$worksheet->write($a-1, $i+3, "GM", $subhead);
	$worksheet->write($a-1, $i+4, "GM%", $subhead);
	$worksheet->write($a-1, $i+5, "GM", $subhead);
	$worksheet->write($a-1, $i+6, "GM%", $subhead);
	
	$worksheet->set_column( $i, $i+1, 8 );
	$worksheet->set_column( $i+2, $i+2, 7 );
	$worksheet->set_column( $i+3, $i+6, undef, undef, 1 );
}

foreach my $i ( 28, 31, 34, 37, 40, 43, 46, 49, 52, 55, 58, 61, 64, 81, 84, 87, 90, 93, 96, 120, 123, 126, 129, 132, 135, 138, 141, 144, 147, 150, 153, 156, 173, 176, 179, 182, 185, 188 ) {
	$worksheet->write($a-1, $i, 'Actual', $subhead);
	$worksheet->write($a-1, $i+1, 'Budget', $subhead);
	$worksheet->write($a-1, $i+2, 'Var', $subhead);
	
	$worksheet->set_column( $i, $i, 8 );
	$worksheet->set_column( $i+1, $i+2, undef, undef, 1 );
}

$worksheet->merge_range( $a-2, 28, $a-2, 30, 'Cost of Sales - Trade', $subhead );
$worksheet->merge_range( $a-2, 31, $a-2, 33, 'Transfer Discrepancy', $subhead );
$worksheet->merge_range( $a-2, 34, $a-2, 36, 'Promotional Item Charged to Margin', $subhead );
$worksheet->merge_range( $a-2, 37, $a-2, 39, 'Wastage', $subhead );
$worksheet->merge_range( $a-2, 40, $a-2, 42, 'Invoice Price Variance - Trade', $subhead );
$worksheet->merge_range( $a-2, 43, $a-2, 45, 'Shrinkage Cost', $subhead );
$worksheet->merge_range( $a-2, 46, $a-2, 48, 'Cost Variance', $subhead );
$worksheet->merge_range( $a-2, 49, $a-2, 51, 'Synchronization Account', $subhead );
$worksheet->merge_range( $a-2, 52, $a-2, 54, 'Freight Recovery', $subhead );
$worksheet->merge_range( $a-2, 55, $a-2, 57, 'Purchase Allowance', $subhead );
$worksheet->merge_range( $a-2, 58, $a-2, 60, 'Purchase Discouns-Special', $subhead );
$worksheet->merge_range( $a-2, 61, $a-2, 63, 'Other Income-Metro Vendor Portal', $subhead );
$worksheet->merge_range( $a-2, 64, $a-2, 66, 'Freight', $subhead );
$worksheet->merge_range( $a-2, 81, $a-2, 83, 'Cost of Sales - Conc', $subhead );
$worksheet->merge_range( $a-2, 84, $a-2, 86, 'Ad Support', $subhead );
$worksheet->merge_range( $a-2, 87, $a-2, 89, 'Other Income-Storage/Concession', $subhead );
$worksheet->merge_range( $a-2, 90, $a-2, 92, 'Light Recovery', $subhead );
$worksheet->merge_range( $a-2, 93, $a-2, 95, 'Water Recovery', $subhead );
$worksheet->merge_range( $a-2, 96, $a-2, 98, 'Supplies Recovery', $subhead );

$worksheet->merge_range( $a-2, 120, $a-2, 122, 'Cost of Sales - Trade', $subhead );
$worksheet->merge_range( $a-2, 123, $a-2, 125, 'Transfer Discrepancy', $subhead );
$worksheet->merge_range( $a-2, 126, $a-2, 128, 'Promotional Item Charged to Margin', $subhead );
$worksheet->merge_range( $a-2, 129, $a-2, 131, 'Wastage', $subhead );
$worksheet->merge_range( $a-2, 132, $a-2, 134, 'Invoice Price Variance - Trade', $subhead );
$worksheet->merge_range( $a-2, 135, $a-2, 137, 'Shrinkage Cost', $subhead );
$worksheet->merge_range( $a-2, 138, $a-2, 140, 'Cost Variance', $subhead );
$worksheet->merge_range( $a-2, 141, $a-2, 143, 'Synchronization Account', $subhead );
$worksheet->merge_range( $a-2, 144, $a-2, 146, 'Freight Recovery', $subhead );
$worksheet->merge_range( $a-2, 147, $a-2, 149, 'Purchase Allowance', $subhead );
$worksheet->merge_range( $a-2, 150, $a-2, 152, 'Purchase Discouns-Special', $subhead );
$worksheet->merge_range( $a-2, 153, $a-2, 155, 'Other Income-Metro Vendor Portal', $subhead );
$worksheet->merge_range( $a-2, 156, $a-2, 158, 'Freight', $subhead );
$worksheet->merge_range( $a-2, 173, $a-2, 175, 'Cost of Sales - Conc', $subhead );
$worksheet->merge_range( $a-2, 176, $a-2, 178, 'Ad Support', $subhead );
$worksheet->merge_range( $a-2, 179, $a-2, 181, 'Other Income-Storage/Concession', $subhead );
$worksheet->merge_range( $a-2, 182, $a-2, 184, 'Light Recovery', $subhead );
$worksheet->merge_range( $a-2, 185, $a-2, 187, 'Water Recovery', $subhead );
$worksheet->merge_range( $a-2, 188, $a-2, 190, 'Supplies Recovery', $subhead );

$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 23 );

}

#sheet 1
sub query_summary{

$sls = $dbh->prepare (qq{
	SELECT SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
			FROM METRO_IT_MARGIN_DEPT
			WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) AND MERCH_GROUP_CODE <> 'Z_OT' 
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND STORE_FORMAT < 5
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
		SELECT A.FORMAT FORMAT, B.O_COST, O_RETAIL, O_MARGIN, C_COST, C_RETAIL, C_MARGIN, AMT432000, AMT433000, AMT458490, AMT434000, AMT458550, AMT460100, AMT460200, AMT460300 , AMT503200 , AMT503250, AMT503500, AMT506000, AMT501000, AMT503000, AMT507000, AMT999998, AMT505000, AMT504000, AMT502000 FROM
			(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_MARGIN_DEPT WHERE STORE_FORMAT <> 3 AND STORE_FORMAT < 5)A LEFT JOIN
			(SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
				FROM METRO_IT_MARGIN_DEPT
				WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) AND MERCH_GROUP_CODE <> 'Z_OT' 
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND STORE_FORMAT < 5
					AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY UPPER(STORE_FORMAT_DESC))B ON A.FORMAT = B.FORMAT 
		ORDER BY 1
		});
	$sls1->execute();
						
		while(my $s = $sls1->fetchrow_hashref()){
								
			$worksheet->merge_range( $a, 4, $a, 6, $s->{FORMAT}, $desc );
						
			$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$border1);
			$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$border1);
				if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
					$worksheet->write($a,9, "",$subt); }
				else{
					$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$subt); }
						
			$worksheet->write($a,10, "",$border1);
			$worksheet->write($a,11, "",$subt);
			$worksheet->write($a,12, "",$border1);
			$worksheet->write($a,13, "",$subt);
			
			$worksheet->write($a,14, $s->{O_RETAIL},$border1);
			$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$border1);
				if ($s->{O_RETAIL} le 0){
					$worksheet->write($a,16, "",$subt); }
				else{
					$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$subt); }
						
			$worksheet->write($a,17, "",$border1);
			$worksheet->write($a,18, "",$subt);
			$worksheet->write($a,19, "",$border1);
			$worksheet->write($a,20, "",$subt);			
			
			$worksheet->write($a,21, $s->{O_RETAIL},$border1);
			$worksheet->write($a,22, $s->{O_MARGIN},$border1);
				if ($s->{O_RETAIL} le 0){
					$worksheet->write($a,23, "",$subt); }
				else{
					$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$subt); }
					
			$worksheet->write($a,24, "",$border1);
			$worksheet->write($a,25, "",$subt);
			$worksheet->write($a,26, "",$border1);
			$worksheet->write($a,27, "",$subt);	
						
			$worksheet->write($a,28, $s->{AMT501000},$border1);
			$worksheet->write($a,29, "",$border1);
			$worksheet->write($a,30, "",$border1);
			
			$worksheet->write($a,31, $s->{AMT503200},$border1);
			$worksheet->write($a,32, "",$border1);
			$worksheet->write($a,33, "",$border1);
						
			$worksheet->write($a,34, $s->{AMT503250},$border1);
			$worksheet->write($a,35, "",$border1);
			$worksheet->write($a,36, "",$border1);
					
			$worksheet->write($a,37, $s->{AMT503500},$border1);
			$worksheet->write($a,38, "",$border1);
			$worksheet->write($a,39, "",$border1);
						
			$worksheet->write($a,40, $s->{AMT506000},$border1);
			$worksheet->write($a,41, "",$border1);
			$worksheet->write($a,42, "",$border1);
						
			$worksheet->write($a,43, $s->{AMT503000},$border1);
			$worksheet->write($a,44, "",$border1);
			$worksheet->write($a,45, "",$border1);				
								
			$worksheet->write($a,46, $s->{AMT507000},$border1);
			$worksheet->write($a,47, "",$border1);
			$worksheet->write($a,48, "",$border1);
				
			$worksheet->write($a,49, $s->{AMT999998},$border1);
			$worksheet->write($a,50, "",$border1);
			$worksheet->write($a,51, "",$border1);
			
			$worksheet->write($a,52, $s->{AMT504000},$border1);
			$worksheet->write($a,53, "",$border1);
			$worksheet->write($a,54, "",$border1);
									
			$worksheet->write($a,55, $s->{AMT432000},$border1);
			$worksheet->write($a,56, "",$border1);
			$worksheet->write($a,57, "",$border1);
						
			$worksheet->write($a,58, $s->{AMT433000},$border1);
			$worksheet->write($a,59, "",$border1);
			$worksheet->write($a,60, "",$border1);
						
			$worksheet->write($a,61, $s->{AMT458490},$border1);
			$worksheet->write($a,62, "",$border1);
			$worksheet->write($a,63, "",$border1);
				
			$worksheet->write($a,64, $s->{AMT505000},$border1);
			$worksheet->write($a,65, "",$border1);
			$worksheet->write($a,66, "",$border1);
						
			$worksheet->write($a,67, $s->{C_RETAIL},$border1);
			$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$border1);
				if ($s->{C_RETAIL} le 0){
					$worksheet->write($a,69, "",$subt); }
				else{
					$worksheet->write($a,69,($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$subt); }
								
			$worksheet->write($a,70, "",$border1);
			$worksheet->write($a,71, "",$subt);
			$worksheet->write($a,72, "",$border1);
			$worksheet->write($a,73, "",$subt);	
				
			$worksheet->write($a,74, $s->{C_RETAIL},$border1);
			$worksheet->write($a,75, $s->{C_MARGIN},$border1);
				if ($s->{C_RETAIL} le 0){
					$worksheet->write($a,76, "",$subt); }
				else{
					$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$subt); }
							
			$worksheet->write($a,77, "",$border1);
			$worksheet->write($a,78, "",$subt);
			$worksheet->write($a,79, "",$border1);
			$worksheet->write($a,80, "",$subt);	
						
			$worksheet->write($a,81, $s->{AMT502000},$border1);
			$worksheet->write($a,82, "",$border1);
			$worksheet->write($a,83, "",$border1);
					
			$worksheet->write($a,84, $s->{AMT434000},$border1);
			$worksheet->write($a,85, "",$border1);
			$worksheet->write($a,86, "",$border1);
								
			$worksheet->write($a,87, $s->{AMT458550},$border1);
			$worksheet->write($a,88, "",$border1);
			$worksheet->write($a,89, "",$border1);
						
			$worksheet->write($a,90, $s->{AMT460100},$border1);
			$worksheet->write($a,91, "",$border1);
			$worksheet->write($a,92, "",$border1);
							
			$worksheet->write($a,93, $s->{AMT460200},$border1);
			$worksheet->write($a,94, "",$border1);
			$worksheet->write($a,95, "",$border1);
								
			$worksheet->write($a,96, $s->{AMT460300},$border1);
			$worksheet->write($a,97, "",$border1);
			$worksheet->write($a,98, "",$border1);	
												
			$a++;
			$counter++;
					
	}
	
	$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$bodyNum);
	$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
		if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
			$worksheet->write($a,9, "",$bodyPct); }
		else{
			$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$bodyPct); }
				
	$worksheet->write($a,10, "",$bodyNum);
	$worksheet->write($a,11, "",$bodyPct);
	$worksheet->write($a,12, "",$bodyNum);
	$worksheet->write($a,13, "",$bodyPct);
					
	$worksheet->write($a,14, $s->{O_RETAIL},$bodyNum);
	$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$bodyNum);
		if ($s->{O_RETAIL} le 0){
			$worksheet->write($a,16, "",$bodyPct); }
		else{
			$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$bodyPct); }
					
	$worksheet->write($a,17, "",$bodyNum);
	$worksheet->write($a,18, "",$bodyPct);
	$worksheet->write($a,19, "",$bodyNum);
	$worksheet->write($a,20, "",$bodyPct);			
					
	$worksheet->write($a,21, $s->{O_RETAIL},$bodyNum);
	$worksheet->write($a,22, $s->{O_MARGIN},$bodyNum);
		if ($s->{O_RETAIL} le 0){
			$worksheet->write($a,23, "",$bodyPct); }
		else{
			$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$bodyPct); }
							
	$worksheet->write($a,24, "",$bodyNum);
	$worksheet->write($a,25, "",$bodyPct);
	$worksheet->write($a,26, "",$bodyNum);
	$worksheet->write($a,27, "",$bodyPct);	
					
	$worksheet->write($a,28, $s->{AMT501000},$bodyNum);
	$worksheet->write($a,29, "",$bodyNum);
	$worksheet->write($a,30, "",$bodyNum);
		
	$worksheet->write($a,31, $s->{AMT503200},$bodyNum);
	$worksheet->write($a,32, "",$bodyNum);
	$worksheet->write($a,33, "",$bodyNum);
		
	$worksheet->write($a,34, $s->{AMT503250},$bodyNum);
	$worksheet->write($a,35, "",$bodyNum);
	$worksheet->write($a,36, "",$bodyNum);
					
	$worksheet->write($a,37, $s->{AMT503500},$bodyNum);
	$worksheet->write($a,38, "",$bodyNum);
	$worksheet->write($a,39, "",$bodyNum);
			
	$worksheet->write($a,40, $s->{AMT506000},$bodyNum);
	$worksheet->write($a,41, "",$bodyNum);
	$worksheet->write($a,42, "",$bodyNum);
					
	$worksheet->write($a,43, $s->{AMT503000},$bodyNum);
	$worksheet->write($a,44, "",$bodyNum);
	$worksheet->write($a,45, "",$bodyNum);				
							
	$worksheet->write($a,46, $s->{AMT507000},$bodyNum);
	$worksheet->write($a,47, "",$bodyNum);
	$worksheet->write($a,48, "",$bodyNum);
					
	$worksheet->write($a,49, $s->{AMT999998},$bodyNum);
	$worksheet->write($a,50, "",$bodyNum);
	$worksheet->write($a,51, "",$bodyNum);
			
	$worksheet->write($a,52, $s->{AMT504000},$bodyNum);
	$worksheet->write($a,53, "",$bodyNum);
	$worksheet->write($a,54, "",$bodyNum);
								
	$worksheet->write($a,55, $s->{AMT432000},$bodyNum);
	$worksheet->write($a,56, "",$bodyNum);
	$worksheet->write($a,57, "",$bodyNum);
					
	$worksheet->write($a,58, $s->{AMT433000},$bodyNum);
	$worksheet->write($a,59, "",$bodyNum);
	$worksheet->write($a,60, "",$bodyNum);
			
	$worksheet->write($a,61, $s->{AMT458490},$bodyNum);
	$worksheet->write($a,62, "",$bodyNum);
	$worksheet->write($a,63, "",$bodyNum);
					
	$worksheet->write($a,64, $s->{AMT505000},$bodyNum);
	$worksheet->write($a,65, "",$bodyNum);
	$worksheet->write($a,66, "",$bodyNum);
					
	$worksheet->write($a,67, $s->{C_RETAIL},$bodyNum);
	$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
		if ($s->{C_RETAIL} le 0){
			$worksheet->write($a,69, "",$bodyPct); }
		else{
			$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$bodyPct); }
							
	$worksheet->write($a,70, "",$bodyNum);
	$worksheet->write($a,71, "",$bodyPct);
	$worksheet->write($a,72, "",$bodyNum);
	$worksheet->write($a,73, "",$bodyPct);	
					
	$worksheet->write($a,74, $s->{C_RETAIL},$bodyNum);
	$worksheet->write($a,75, $s->{C_MARGIN},$bodyNum);
		if ($s->{C_RETAIL} le 0){
			$worksheet->write($a,76, "",$bodyPct); }
		else{
			$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$bodyPct); }
						
	$worksheet->write($a,77, "",$bodyNum);
	$worksheet->write($a,78, "",$bodyPct);
	$worksheet->write($a,79, "",$bodyNum);
	$worksheet->write($a,80, "",$bodyPct);	
		
	$worksheet->write($a,81, $s->{AMT502000},$bodyNum);
	$worksheet->write($a,82, "",$bodyNum);
	$worksheet->write($a,83, "",$bodyNum);
							
	$worksheet->write($a,84, $s->{AMT434000},$bodyNum);
	$worksheet->write($a,85, "",$bodyNum);
	$worksheet->write($a,86, "",$bodyNum);
							
	$worksheet->write($a,87, $s->{AMT458550},$bodyNum);
	$worksheet->write($a,88, "",$bodyNum);
	$worksheet->write($a,89, "",$bodyNum);
					
	$worksheet->write($a,90, $s->{AMT460100},$bodyNum);
	$worksheet->write($a,91, "",$bodyNum);
	$worksheet->write($a,92, "",$bodyNum);
						
	$worksheet->write($a,93, $s->{AMT460200},$bodyNum);
	$worksheet->write($a,94, "",$bodyNum);
	$worksheet->write($a,95, "",$bodyNum);
							
	$worksheet->write($a,96, $s->{AMT460300},$bodyNum);
	$worksheet->write($a,97, "",$bodyNum);
	$worksheet->write($a,98, "",$bodyNum);
	
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
	SELECT SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
			FROM METRO_IT_MARGIN_DEPT
			WHERE MERCH_GROUP_CODE <> 'Z_OT' 
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
				AND UPPER(STORE_FORMAT_DESC) = '$store_format' AND STORE_FORMAT < 5
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
		SELECT MATURED_FLG, CASE WHEN MATURED_FLG = 1 THEN 'COMP' ELSE 'NEW' END AS FLG_DESC, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
		FROM METRO_IT_MARGIN_DEPT
		WHERE MERCH_GROUP_CODE <> 'Z_OT'
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
			AND UPPER(STORE_FORMAT_DESC) = '$store_format' AND STORE_FORMAT < 5
			AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
		GROUP BY MATURED_FLG, CASE WHEN MATURED_FLG = 1 THEN 'COMP' ELSE 'NEW' END
		ORDER BY MATURED_FLG DESC
		});
	$sls1->execute();
		
		$format_counter = $a;
		while(my $s = $sls1->fetchrow_hashref()){
		$flg = $s->{MATURED_FLG};
		$flg_desc = $s->{FLG_DESC};

		$sls2 = $dbh->prepare (qq{
			SELECT STORE_CODE, STORE_DESCRIPTION, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
			FROM METRO_IT_MARGIN_DEPT
			WHERE MERCH_GROUP_CODE <> 'Z_OT' AND MATURED_FLG = '$flg' 
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
				AND UPPER(STORE_FORMAT_DESC) = '$store_format'  AND STORE_FORMAT < 5 
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY STORE_CODE, STORE_DESCRIPTION
			ORDER BY 1
			});
		$sls2->execute();
			
			while(my $s = $sls2->fetchrow_hashref()){
									
				$worksheet->write( $a, 5, $s->{STORE_CODE}, $desc );
				$worksheet->write( $a, 6, $s->{STORE_DESCRIPTION}, $desc );
				$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$border1);
				$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$border1);
					if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
						$worksheet->write($a,9, "",$subt); }
					else{
						$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$subt); }
							
				$worksheet->write($a,10, "",$border1);
				$worksheet->write($a,11, "",$subt);
				$worksheet->write($a,12, "",$border1);
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{O_RETAIL},$border1);
				$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$border1);
					if ($s->{O_RETAIL} le 0){
						$worksheet->write($a,16, "",$subt); }
					else{
						$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$subt); }
							
				$worksheet->write($a,17, "",$border1);
				$worksheet->write($a,18, "",$subt);
				$worksheet->write($a,19, "",$border1);
				$worksheet->write($a,20, "",$subt);			
				
				$worksheet->write($a,21, $s->{O_RETAIL},$border1);
				$worksheet->write($a,22, $s->{O_MARGIN},$border1);
					if ($s->{O_RETAIL} le 0){
						$worksheet->write($a,23, "",$subt); }
					else{
						$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$subt); }
						
				$worksheet->write($a,24, "",$border1);
				$worksheet->write($a,25, "",$subt);
				$worksheet->write($a,26, "",$border1);
				$worksheet->write($a,27, "",$subt);	
							
				$worksheet->write($a,28, $s->{AMT501000},$border1);
				$worksheet->write($a,29, "",$border1);
				$worksheet->write($a,30, "",$border1);
				
				$worksheet->write($a,31, $s->{AMT503200},$border1);
				$worksheet->write($a,32, "",$border1);
				$worksheet->write($a,33, "",$border1);
							
				$worksheet->write($a,34, $s->{AMT503250},$border1);
				$worksheet->write($a,35, "",$border1);
				$worksheet->write($a,36, "",$border1);
						
				$worksheet->write($a,37, $s->{AMT503500},$border1);
				$worksheet->write($a,38, "",$border1);
				$worksheet->write($a,39, "",$border1);
							
				$worksheet->write($a,40, $s->{AMT506000},$border1);
				$worksheet->write($a,41, "",$border1);
				$worksheet->write($a,42, "",$border1);
							
				$worksheet->write($a,43, $s->{AMT503000},$border1);
				$worksheet->write($a,44, "",$border1);
				$worksheet->write($a,45, "",$border1);				
									
				$worksheet->write($a,46, $s->{AMT507000},$border1);
				$worksheet->write($a,47, "",$border1);
				$worksheet->write($a,48, "",$border1);
					
				$worksheet->write($a,49, $s->{AMT999998},$border1);
				$worksheet->write($a,50, "",$border1);
				$worksheet->write($a,51, "",$border1);
				
				$worksheet->write($a,52, $s->{AMT504000},$border1);
				$worksheet->write($a,53, "",$border1);
				$worksheet->write($a,54, "",$border1);
										
				$worksheet->write($a,55, $s->{AMT432000},$border1);
				$worksheet->write($a,56, "",$border1);
				$worksheet->write($a,57, "",$border1);
							
				$worksheet->write($a,58, $s->{AMT433000},$border1);
				$worksheet->write($a,59, "",$border1);
				$worksheet->write($a,60, "",$border1);
							
				$worksheet->write($a,61, $s->{AMT458490},$border1);
				$worksheet->write($a,62, "",$border1);
				$worksheet->write($a,63, "",$border1);
					
				$worksheet->write($a,64, $s->{AMT505000},$border1);
				$worksheet->write($a,65, "",$border1);
				$worksheet->write($a,66, "",$border1);
							
				$worksheet->write($a,67, $s->{C_RETAIL},$border1);
				$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$border1);
					if ($s->{C_RETAIL} le 0){
						$worksheet->write($a,69, "",$subt); }
					else{
						$worksheet->write($a,69,($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$subt); }
									
				$worksheet->write($a,70, "",$border1);
				$worksheet->write($a,71, "",$subt);
				$worksheet->write($a,72, "",$border1);
				$worksheet->write($a,73, "",$subt);	
					
				$worksheet->write($a,74, $s->{C_RETAIL},$border1);
				$worksheet->write($a,75, $s->{C_MARGIN},$border1);
					if ($s->{C_RETAIL} le 0){
						$worksheet->write($a,76, "",$subt); }
					else{
						$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$subt); }
								
				$worksheet->write($a,77, "",$border1);
				$worksheet->write($a,78, "",$subt);
				$worksheet->write($a,79, "",$border1);
				$worksheet->write($a,80, "",$subt);	
							
				$worksheet->write($a,81, $s->{AMT502000},$border1);
				$worksheet->write($a,82, "",$border1);
				$worksheet->write($a,83, "",$border1);
						
				$worksheet->write($a,84, $s->{AMT434000},$border1);
				$worksheet->write($a,85, "",$border1);
				$worksheet->write($a,86, "",$border1);
									
				$worksheet->write($a,87, $s->{AMT458550},$border1);
				$worksheet->write($a,88, "",$border1);
				$worksheet->write($a,89, "",$border1);
							
				$worksheet->write($a,90, $s->{AMT460100},$border1);
				$worksheet->write($a,91, "",$border1);
				$worksheet->write($a,92, "",$border1);
								
				$worksheet->write($a,93, $s->{AMT460200},$border1);
				$worksheet->write($a,94, "",$border1);
				$worksheet->write($a,95, "",$border1);
									
				$worksheet->write($a,96, $s->{AMT460300},$border1);
				$worksheet->write($a,97, "",$border1);
				$worksheet->write($a,98, "",$border1);	
													
			$a++;
			$counter++;
						
		}
		
		$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$bodyNum);
		$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
			if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
				$worksheet->write($a,9, "",$bodyPct); }
			else{
				$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$bodyPct); }
					
		$worksheet->write($a,10, "",$bodyNum);
		$worksheet->write($a,11, "",$bodyPct);
		$worksheet->write($a,12, "",$bodyNum);
		$worksheet->write($a,13, "",$bodyPct);
						
		$worksheet->write($a,14, $s->{O_RETAIL},$bodyNum);
		$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$bodyNum);
			if ($s->{O_RETAIL} le 0){
				$worksheet->write($a,16, "",$bodyPct); }
			else{
				$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$bodyPct); }
						
		$worksheet->write($a,17, "",$bodyNum);
		$worksheet->write($a,18, "",$bodyPct);
		$worksheet->write($a,19, "",$bodyNum);
		$worksheet->write($a,20, "",$bodyPct);			
						
		$worksheet->write($a,21, $s->{O_RETAIL},$bodyNum);
		$worksheet->write($a,22, $s->{O_MARGIN},$bodyNum);
			if ($s->{O_RETAIL} le 0){
				$worksheet->write($a,23, "",$bodyPct); }
			else{
				$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$bodyPct); }
								
		$worksheet->write($a,24, "",$bodyNum);
		$worksheet->write($a,25, "",$bodyPct);
		$worksheet->write($a,26, "",$bodyNum);
		$worksheet->write($a,27, "",$bodyPct);	
						
		$worksheet->write($a,28, $s->{AMT501000},$bodyNum);
		$worksheet->write($a,29, "",$bodyNum);
		$worksheet->write($a,30, "",$bodyNum);
			
		$worksheet->write($a,31, $s->{AMT503200},$bodyNum);
		$worksheet->write($a,32, "",$bodyNum);
		$worksheet->write($a,33, "",$bodyNum);
			
		$worksheet->write($a,34, $s->{AMT503250},$bodyNum);
		$worksheet->write($a,35, "",$bodyNum);
		$worksheet->write($a,36, "",$bodyNum);
						
		$worksheet->write($a,37, $s->{AMT503500},$bodyNum);
		$worksheet->write($a,38, "",$bodyNum);
		$worksheet->write($a,39, "",$bodyNum);
				
		$worksheet->write($a,40, $s->{AMT506000},$bodyNum);
		$worksheet->write($a,41, "",$bodyNum);
		$worksheet->write($a,42, "",$bodyNum);
						
		$worksheet->write($a,43, $s->{AMT503000},$bodyNum);
		$worksheet->write($a,44, "",$bodyNum);
		$worksheet->write($a,45, "",$bodyNum);				
								
		$worksheet->write($a,46, $s->{AMT507000},$bodyNum);
		$worksheet->write($a,47, "",$bodyNum);
		$worksheet->write($a,48, "",$bodyNum);
						
		$worksheet->write($a,49, $s->{AMT999998},$bodyNum);
		$worksheet->write($a,50, "",$bodyNum);
		$worksheet->write($a,51, "",$bodyNum);
				
		$worksheet->write($a,52, $s->{AMT504000},$bodyNum);
		$worksheet->write($a,53, "",$bodyNum);
		$worksheet->write($a,54, "",$bodyNum);
									
		$worksheet->write($a,55, $s->{AMT432000},$bodyNum);
		$worksheet->write($a,56, "",$bodyNum);
		$worksheet->write($a,57, "",$bodyNum);
						
		$worksheet->write($a,58, $s->{AMT433000},$bodyNum);
		$worksheet->write($a,59, "",$bodyNum);
		$worksheet->write($a,60, "",$bodyNum);
				
		$worksheet->write($a,61, $s->{AMT458490},$bodyNum);
		$worksheet->write($a,62, "",$bodyNum);
		$worksheet->write($a,63, "",$bodyNum);
						
		$worksheet->write($a,64, $s->{AMT505000},$bodyNum);
		$worksheet->write($a,65, "",$bodyNum);
		$worksheet->write($a,66, "",$bodyNum);
						
		$worksheet->write($a,67, $s->{C_RETAIL},$bodyNum);
		$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
			if ($s->{C_RETAIL} le 0){
				$worksheet->write($a,69, "",$bodyPct); }
			else{
				$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$bodyPct); }
								
		$worksheet->write($a,70, "",$bodyNum);
		$worksheet->write($a,71, "",$bodyPct);
		$worksheet->write($a,72, "",$bodyNum);
		$worksheet->write($a,73, "",$bodyPct);	
						
		$worksheet->write($a,74, $s->{C_RETAIL},$bodyNum);
		$worksheet->write($a,75, $s->{C_MARGIN},$bodyNum);
			if ($s->{C_RETAIL} le 0){
				$worksheet->write($a,76, "",$bodyPct); }
			else{
				$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$bodyPct); }
							
		$worksheet->write($a,77, "",$bodyNum);
		$worksheet->write($a,78, "",$bodyPct);
		$worksheet->write($a,79, "",$bodyNum);
		$worksheet->write($a,80, "",$bodyPct);	
			
		$worksheet->write($a,81, $s->{AMT502000},$bodyNum);
		$worksheet->write($a,82, "",$bodyNum);
		$worksheet->write($a,83, "",$bodyNum);
								
		$worksheet->write($a,84, $s->{AMT434000},$bodyNum);
		$worksheet->write($a,85, "",$bodyNum);
		$worksheet->write($a,86, "",$bodyNum);
								
		$worksheet->write($a,87, $s->{AMT458550},$bodyNum);
		$worksheet->write($a,88, "",$bodyNum);
		$worksheet->write($a,89, "",$bodyNum);
						
		$worksheet->write($a,90, $s->{AMT460100},$bodyNum);
		$worksheet->write($a,91, "",$bodyNum);
		$worksheet->write($a,92, "",$bodyNum);
							
		$worksheet->write($a,93, $s->{AMT460200},$bodyNum);
		$worksheet->write($a,94, "",$bodyNum);
		$worksheet->write($a,95, "",$bodyNum);
								
		$worksheet->write($a,96, $s->{AMT460300},$bodyNum);
		$worksheet->write($a,97, "",$bodyNum);
		$worksheet->write($a,98, "",$bodyNum);
		
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $flg_desc, $border2 );
		
		$a++;
		$counter = 0;
	}
	
	$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$bodyNum);
	$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
		if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
			$worksheet->write($a,9, "",$bodyPct); }
		else{
			$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$bodyPct); }
				
	$worksheet->write($a,10, "",$bodyNum);
	$worksheet->write($a,11, "",$bodyPct);
	$worksheet->write($a,12, "",$bodyNum);
	$worksheet->write($a,13, "",$bodyPct);
					
	$worksheet->write($a,14, $s->{O_RETAIL},$bodyNum);
	$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$bodyNum);
		if ($s->{O_RETAIL} le 0){
			$worksheet->write($a,16, "",$bodyPct); }
		else{
			$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$bodyPct); }
					
	$worksheet->write($a,17, "",$bodyNum);
	$worksheet->write($a,18, "",$bodyPct);
	$worksheet->write($a,19, "",$bodyNum);
	$worksheet->write($a,20, "",$bodyPct);			
					
	$worksheet->write($a,21, $s->{O_RETAIL},$bodyNum);
	$worksheet->write($a,22, $s->{O_MARGIN},$bodyNum);
		if ($s->{O_RETAIL} le 0){
			$worksheet->write($a,23, "",$bodyPct); }
		else{
			$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$bodyPct); }
							
	$worksheet->write($a,24, "",$bodyNum);
	$worksheet->write($a,25, "",$bodyPct);
	$worksheet->write($a,26, "",$bodyNum);
	$worksheet->write($a,27, "",$bodyPct);	
					
	$worksheet->write($a,28, $s->{AMT501000},$bodyNum);
	$worksheet->write($a,29, "",$bodyNum);
	$worksheet->write($a,30, "",$bodyNum);
		
	$worksheet->write($a,31, $s->{AMT503200},$bodyNum);
	$worksheet->write($a,32, "",$bodyNum);
	$worksheet->write($a,33, "",$bodyNum);
		
	$worksheet->write($a,34, $s->{AMT503250},$bodyNum);
	$worksheet->write($a,35, "",$bodyNum);
	$worksheet->write($a,36, "",$bodyNum);
					
	$worksheet->write($a,37, $s->{AMT503500},$bodyNum);
	$worksheet->write($a,38, "",$bodyNum);
	$worksheet->write($a,39, "",$bodyNum);
			
	$worksheet->write($a,40, $s->{AMT506000},$bodyNum);
	$worksheet->write($a,41, "",$bodyNum);
	$worksheet->write($a,42, "",$bodyNum);
					
	$worksheet->write($a,43, $s->{AMT503000},$bodyNum);
	$worksheet->write($a,44, "",$bodyNum);
	$worksheet->write($a,45, "",$bodyNum);				
							
	$worksheet->write($a,46, $s->{AMT507000},$bodyNum);
	$worksheet->write($a,47, "",$bodyNum);
	$worksheet->write($a,48, "",$bodyNum);
					
	$worksheet->write($a,49, $s->{AMT999998},$bodyNum);
	$worksheet->write($a,50, "",$bodyNum);
	$worksheet->write($a,51, "",$bodyNum);
			
	$worksheet->write($a,52, $s->{AMT504000},$bodyNum);
	$worksheet->write($a,53, "",$bodyNum);
	$worksheet->write($a,54, "",$bodyNum);
								
	$worksheet->write($a,55, $s->{AMT432000},$bodyNum);
	$worksheet->write($a,56, "",$bodyNum);
	$worksheet->write($a,57, "",$bodyNum);
					
	$worksheet->write($a,58, $s->{AMT433000},$bodyNum);
	$worksheet->write($a,59, "",$bodyNum);
	$worksheet->write($a,60, "",$bodyNum);
			
	$worksheet->write($a,61, $s->{AMT458490},$bodyNum);
	$worksheet->write($a,62, "",$bodyNum);
	$worksheet->write($a,63, "",$bodyNum);
					
	$worksheet->write($a,64, $s->{AMT505000},$bodyNum);
	$worksheet->write($a,65, "",$bodyNum);
	$worksheet->write($a,66, "",$bodyNum);
					
	$worksheet->write($a,67, $s->{C_RETAIL},$bodyNum);
	$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
		if ($s->{C_RETAIL} le 0){
			$worksheet->write($a,69, "",$bodyPct); }
		else{
			$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$bodyPct); }
							
	$worksheet->write($a,70, "",$bodyNum);
	$worksheet->write($a,71, "",$bodyPct);
	$worksheet->write($a,72, "",$bodyNum);
	$worksheet->write($a,73, "",$bodyPct);	
					
	$worksheet->write($a,74, $s->{C_RETAIL},$bodyNum);
	$worksheet->write($a,75, $s->{C_MARGIN},$bodyNum);
		if ($s->{C_RETAIL} le 0){
			$worksheet->write($a,76, "",$bodyPct); }
		else{
			$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$bodyPct); }
						
	$worksheet->write($a,77, "",$bodyNum);
	$worksheet->write($a,78, "",$bodyPct);
	$worksheet->write($a,79, "",$bodyNum);
	$worksheet->write($a,80, "",$bodyPct);	
		
	$worksheet->write($a,81, $s->{AMT502000},$bodyNum);
	$worksheet->write($a,82, "",$bodyNum);
	$worksheet->write($a,83, "",$bodyNum);
							
	$worksheet->write($a,84, $s->{AMT434000},$bodyNum);
	$worksheet->write($a,85, "",$bodyNum);
	$worksheet->write($a,86, "",$bodyNum);
							
	$worksheet->write($a,87, $s->{AMT458550},$bodyNum);
	$worksheet->write($a,88, "",$bodyNum);
	$worksheet->write($a,89, "",$bodyNum);
					
	$worksheet->write($a,90, $s->{AMT460100},$bodyNum);
	$worksheet->write($a,91, "",$bodyNum);
	$worksheet->write($a,92, "",$bodyNum);
						
	$worksheet->write($a,93, $s->{AMT460200},$bodyNum);
	$worksheet->write($a,94, "",$bodyNum);
	$worksheet->write($a,95, "",$bodyNum);
							
	$worksheet->write($a,96, $s->{AMT460300},$bodyNum);
	$worksheet->write($a,97, "",$bodyNum);
	$worksheet->write($a,98, "",$bodyNum);
	
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
	SELECT SUM(O_COST_DS) AS O_COST_DS, SUM(O_RETAIL_DS) AS O_RETAIL_DS, SUM(O_MARGIN_DS) AS O_MARGIN_DS, SUM(C_COST_DS) AS C_COST_DS, SUM(C_RETAIL_DS) AS C_RETAIL_DS, SUM(C_MARGIN_DS) AS C_MARGIN_DS, SUM(AMT432000_DS) AS AMT432000_DS, SUM(AMT433000_DS) AS AMT433000_DS, SUM(AMT458490_DS) AS AMT458490_DS, SUM(AMT434000_DS) AS AMT434000_DS, SUM(AMT458550_DS) AS AMT458550_DS, SUM(AMT460100_DS) AS AMT460100_DS, SUM(AMT460200_DS) AS AMT460200_DS, SUM(AMT460300_DS) AS AMT460300_DS, SUM(AMT503200_DS) AS AMT503200_DS, SUM(AMT503250_DS) AS AMT503250_DS, SUM(AMT503500_DS) AS AMT503500_DS, SUM(AMT506000_DS) AS AMT506000_DS, SUM(AMT501000_DS) AS AMT501000_DS, SUM(AMT503000_DS) AS AMT503000_DS, SUM(AMT507000_DS) AS AMT507000_DS, SUM(AMT999998_DS) AS AMT999998_DS, SUM(AMT505000_DS) AS AMT505000_DS, SUM(AMT504000_DS) AS AMT504000_DS, SUM(AMT502000_DS) AS AMT502000_DS, SUM(O_COST_SU) AS O_COST_SU, SUM(O_RETAIL_SU) AS O_RETAIL_SU, SUM(O_MARGIN_SU) AS O_MARGIN_SU, SUM(C_COST_SU) AS C_COST_SU, SUM(C_RETAIL_SU) AS C_RETAIL_SU, SUM(C_MARGIN_SU) AS C_MARGIN_SU, SUM(AMT432000_SU) AS AMT432000_SU, SUM(AMT433000_SU) AS AMT433000_SU, SUM(AMT458490_SU) AS AMT458490_SU, SUM(AMT434000_SU) AS AMT434000_SU, SUM(AMT458550_SU) AS AMT458550_SU, SUM(AMT460100_SU) AS AMT460100_SU, SUM(AMT460200_SU) AS AMT460200_SU, SUM(AMT460300_SU) AS AMT460300_SU, SUM(AMT503200_SU) AS AMT503200_SU, SUM(AMT503250_SU) AS AMT503250_SU, SUM(AMT503500_SU) AS AMT503500_SU, SUM(AMT506000_SU) AS AMT506000_SU, SUM(AMT501000_SU) AS AMT501000_SU, SUM(AMT503000_SU) AS AMT503000_SU, SUM(AMT507000_SU) AS AMT507000_SU, SUM(AMT999998_SU) AS AMT999998_SU, SUM(AMT505000_SU) AS AMT505000_SU, SUM(AMT504000_SU) AS AMT504000_SU, SUM(AMT502000_SU) AS AMT502000_SU
	FROM
		(SELECT SUM(OTR_COST) AS O_COST_DS, SUM(OTR_RETAIL) AS O_RETAIL_DS, SUM(OTR_MARGIN) AS O_MARGIN_DS, SUM(CON_COST) AS C_COST_DS, SUM(CON_RETAIL) AS C_RETAIL_DS, SUM(CON_MARGIN) AS C_MARGIN_DS, SUM(AMOUNT432000) AS AMT432000_DS, SUM(AMOUNT433000) AS AMT433000_DS, SUM(AMOUNT458490) AS AMT458490_DS, SUM(AMOUNT434000) AS AMT434000_DS, SUM(AMOUNT458550) AS AMT458550_DS, SUM(AMOUNT460100) AS AMT460100_DS, SUM(AMOUNT460200) AS AMT460200_DS, SUM(AMOUNT460300) AS AMT460300_DS, SUM(AMOUNT503200) AS AMT503200_DS, SUM(AMOUNT503250) AS AMT503250_DS, SUM(AMOUNT503500) AS AMT503500_DS, SUM(AMOUNT506000) AS AMT506000_DS, SUM(AMOUNT501000) AS AMT501000_DS, SUM(AMOUNT503000) AS AMT503000_DS, SUM(AMOUNT507000) AS AMT507000_DS, SUM(AMOUNT999998) AS AMT999998_DS, SUM(AMOUNT505000) AS AMT505000_DS, SUM(AMOUNT504000) AS AMT504000_DS, SUM(AMOUNT502000) AS AMT502000_DS, 0 AS O_COST_SU, 0 AS O_RETAIL_SU, 0 AS O_MARGIN_SU, 0 AS C_COST_SU, 0 AS C_RETAIL_SU, 0 AS C_MARGIN_SU, 0 AS AMT432000_SU, 0 AS AMT433000_SU, 0 AS AMT458490_SU, 0 AS AMT434000_SU, 0 AS AMT458550_SU, 0 AS AMT460100_SU, 0 AS AMT460200_SU, 0 AS AMT460300_SU, 0 AS AMT503200_SU, 0 AS AMT503250_SU, 0 AS AMT503500_SU, 0 AS AMT506000_SU, 0 AS AMT501000_SU, 0 AS AMT503000_SU, 0 AS AMT507000_SU, 0 AS AMT999998_SU, 0 AS AMT505000_SU, 0 AS AMT504000_SU, 0 AS AMT502000_SU
		FROM METRO_IT_MARGIN_DEPT
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) AND MERCH_GROUP_CODE = 'DS' 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND STORE_FORMAT < 5
			AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
		UNION ALL
		SELECT 0 AS O_COST_DS, 0 AS O_RETAIL_DS, 0 AS O_MARGIN_DS, 0 AS C_COST_DS, 0 AS C_RETAIL_DS, 0 AS C_MARGIN_DS, 0 AS AMT432000_DS, 0 AS AMT433000_DS, 0 AS AMT458490_DS, 0 AS AMT434000_DS, 0 AS AMT458550_DS, 0 AS AMT460100_DS, 0 AS AMT460200_DS, 0 AS AMT460300_DS, 0 AS AMT503200_DS, 0 AS AMT503250_DS, 0 AS AMT503500_DS, 0 AS AMT506000_DS, 0 AS AMT501000_DS, 0 AS AMT503000_DS, 0 AS AMT507000_DS, 0 AS AMT999998_DS, 0 AS AMT505000_DS, 0 AS AMT504000_DS, 0 AS AMT502000_DS, SUM(OTR_COST) AS O_COST_SU, SUM(OTR_RETAIL) AS O_RETAIL_SU, SUM(OTR_MARGIN) AS O_MARGIN_SU, SUM(CON_COST) AS C_COST_SU, SUM(CON_RETAIL) AS C_RETAIL_SU, SUM(CON_MARGIN) AS C_MARGIN_SU, SUM(AMOUNT432000) AS AMT432000_SU, SUM(AMOUNT433000) AS AMT433000_SU, SUM(AMOUNT458490) AS AMT458490_SU, SUM(AMOUNT434000) AS AMT434000_SU, SUM(AMOUNT458550) AS AMT458550_SU, SUM(AMOUNT460100) AS AMT460100_SU, SUM(AMOUNT460200) AS AMT460200_SU, SUM(AMOUNT460300) AS AMT460300_SU, SUM(AMOUNT503200) AS AMT503200_SU, SUM(AMOUNT503250) AS AMT503250_SU, SUM(AMOUNT503500) AS AMT503500_SU, SUM(AMOUNT506000) AS AMT506000_SU, SUM(AMOUNT501000) AS AMT501000_SU, SUM(AMOUNT503000) AS AMT503000_SU, SUM(AMOUNT507000) AS AMT507000_SU, SUM(AMOUNT999998) AS AMT999998_SU, SUM(AMOUNT505000) AS AMT505000_SU, SUM(AMOUNT504000) AS AMT504000_SU, SUM(AMOUNT502000) AS AMT502000_SU
		FROM METRO_IT_MARGIN_DEPT
		WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) AND MERCH_GROUP_CODE = 'SU' 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND STORE_FORMAT < 5
			AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') ))
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
		SELECT A.FORMAT, SUM(O_COST_DS) AS O_COST_DS, SUM(O_RETAIL_DS) AS O_RETAIL_DS, SUM(O_MARGIN_DS) AS O_MARGIN_DS, SUM(C_COST_DS) AS C_COST_DS, SUM(C_RETAIL_DS) AS C_RETAIL_DS, SUM(C_MARGIN_DS) AS C_MARGIN_DS, SUM(AMT432000_DS) AS AMT432000_DS, SUM(AMT433000_DS) AS AMT433000_DS, SUM(AMT458490_DS) AS AMT458490_DS, SUM(AMT434000_DS) AS AMT434000_DS, SUM(AMT458550_DS) AS AMT458550_DS, SUM(AMT460100_DS) AS AMT460100_DS, SUM(AMT460200_DS) AS AMT460200_DS, SUM(AMT460300_DS) AS AMT460300_DS, SUM(AMT503200_DS) AS AMT503200_DS, SUM(AMT503250_DS) AS AMT503250_DS, SUM(AMT503500_DS) AS AMT503500_DS, SUM(AMT506000_DS) AS AMT506000_DS, SUM(AMT501000_DS) AS AMT501000_DS, SUM(AMT503000_DS) AS AMT503000_DS, SUM(AMT507000_DS) AS AMT507000_DS, SUM(AMT999998_DS) AS AMT999998_DS, SUM(AMT505000_DS) AS AMT505000_DS, SUM(AMT504000_DS) AS AMT504000_DS, SUM(AMT502000_DS) AS AMT502000_DS, SUM(O_COST_SU) AS O_COST_SU, SUM(O_RETAIL_SU) AS O_RETAIL_SU, SUM(O_MARGIN_SU) AS O_MARGIN_SU, SUM(C_COST_SU) AS C_COST_SU, SUM(C_RETAIL_SU) AS C_RETAIL_SU, SUM(C_MARGIN_SU) AS C_MARGIN_SU, SUM(AMT432000_SU) AS AMT432000_SU, SUM(AMT433000_SU) AS AMT433000_SU, SUM(AMT458490_SU) AS AMT458490_SU, SUM(AMT434000_SU) AS AMT434000_SU, SUM(AMT458550_SU) AS AMT458550_SU, SUM(AMT460100_SU) AS AMT460100_SU, SUM(AMT460200_SU) AS AMT460200_SU, SUM(AMT460300_SU) AS AMT460300_SU, SUM(AMT503200_SU) AS AMT503200_SU, SUM(AMT503250_SU) AS AMT503250_SU, SUM(AMT503500_SU) AS AMT503500_SU, SUM(AMT506000_SU) AS AMT506000_SU, SUM(AMT501000_SU) AS AMT501000_SU, SUM(AMT503000_SU) AS AMT503000_SU, SUM(AMT507000_SU) AS AMT507000_SU, SUM(AMT999998_SU) AS AMT999998_SU, SUM(AMT505000_SU) AS AMT505000_SU, SUM(AMT504000_SU) AS AMT504000_SU, SUM(AMT502000_SU) AS AMT502000_SU
		FROM
			(
			(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_MARGIN_DEPT WHERE STORE_FORMAT <> 3 AND STORE_FORMAT < 5)A LEFT JOIN
			(SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, SUM(OTR_COST) AS O_COST_DS, SUM(OTR_RETAIL) AS O_RETAIL_DS, SUM(OTR_MARGIN) AS O_MARGIN_DS, SUM(CON_COST) AS C_COST_DS, SUM(CON_RETAIL) AS C_RETAIL_DS, SUM(CON_MARGIN) AS C_MARGIN_DS, SUM(AMOUNT432000) AS AMT432000_DS, SUM(AMOUNT433000) AS AMT433000_DS, SUM(AMOUNT458490) AS AMT458490_DS, SUM(AMOUNT434000) AS AMT434000_DS, SUM(AMOUNT458550) AS AMT458550_DS, SUM(AMOUNT460100) AS AMT460100_DS, SUM(AMOUNT460200) AS AMT460200_DS, SUM(AMOUNT460300) AS AMT460300_DS, SUM(AMOUNT503200) AS AMT503200_DS, SUM(AMOUNT503250) AS AMT503250_DS, SUM(AMOUNT503500) AS AMT503500_DS, SUM(AMOUNT506000) AS AMT506000_DS, SUM(AMOUNT501000) AS AMT501000_DS, SUM(AMOUNT503000) AS AMT503000_DS, SUM(AMOUNT507000) AS AMT507000_DS, SUM(AMOUNT999998) AS AMT999998_DS, SUM(AMOUNT505000) AS AMT505000_DS, SUM(AMOUNT504000) AS AMT504000_DS, SUM(AMOUNT502000) AS AMT502000_DS, 0 AS O_COST_SU, 0 AS O_RETAIL_SU, 0 AS O_MARGIN_SU, 0 AS C_COST_SU, 0 AS C_RETAIL_SU, 0 AS C_MARGIN_SU, 0 AS AMT432000_SU, 0 AS AMT433000_SU, 0 AS AMT458490_SU, 0 AS AMT434000_SU, 0 AS AMT458550_SU, 0 AS AMT460100_SU, 0 AS AMT460200_SU, 0 AS AMT460300_SU, 0 AS AMT503200_SU, 0 AS AMT503250_SU, 0 AS AMT503500_SU, 0 AS AMT506000_SU, 0 AS AMT501000_SU, 0 AS AMT503000_SU, 0 AS AMT507000_SU, 0 AS AMT999998_SU, 0 AS AMT505000_SU, 0 AS AMT504000_SU, 0 AS AMT502000_SU
			FROM METRO_IT_MARGIN_DEPT
			WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) AND MERCH_GROUP_CODE = 'DS' 
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND STORE_FORMAT < 5
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY UPPER(STORE_FORMAT_DESC)
			UNION ALL
			SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, 0 AS O_COST_DS, 0 AS O_RETAIL_DS, 0 AS O_MARGIN_DS, 0 AS C_COST_DS, 0 AS C_RETAIL_DS, 0 AS C_MARGIN_DS, 0 AS AMT432000_DS, 0 AS AMT433000_DS, 0 AS AMT458490_DS, 0 AS AMT434000_DS, 0 AS AMT458550_DS, 0 AS AMT460100_DS, 0 AS AMT460200_DS, 0 AS AMT460300_DS, 0 AS AMT503200_DS, 0 AS AMT503250_DS, 0 AS AMT503500_DS, 0 AS AMT506000_DS, 0 AS AMT501000_DS, 0 AS AMT503000_DS, 0 AS AMT507000_DS, 0 AS AMT999998_DS, 0 AS AMT505000_DS, 0 AS AMT504000_DS, 0 AS AMT502000_DS, SUM(OTR_COST) AS O_COST_SU, SUM(OTR_RETAIL) AS O_RETAIL_SU, SUM(OTR_MARGIN) AS O_MARGIN_SU, SUM(CON_COST) AS C_COST_SU, SUM(CON_RETAIL) AS C_RETAIL_SU, SUM(CON_MARGIN) AS C_MARGIN_SU, SUM(AMOUNT432000) AS AMT432000_SU, SUM(AMOUNT433000) AS AMT433000_SU, SUM(AMOUNT458490) AS AMT458490_SU, SUM(AMOUNT434000) AS AMT434000_SU, SUM(AMOUNT458550) AS AMT458550_SU, SUM(AMOUNT460100) AS AMT460100_SU, SUM(AMOUNT460200) AS AMT460200_SU, SUM(AMOUNT460300) AS AMT460300_SU, SUM(AMOUNT503200) AS AMT503200_SU, SUM(AMOUNT503250) AS AMT503250_SU, SUM(AMOUNT503500) AS AMT503500_SU, SUM(AMOUNT506000) AS AMT506000_SU, SUM(AMOUNT501000) AS AMT501000_SU, SUM(AMOUNT503000) AS AMT503000_SU, SUM(AMOUNT507000) AS AMT507000_SU, SUM(AMOUNT999998) AS AMT999998_SU, SUM(AMOUNT505000) AS AMT505000_SU, SUM(AMOUNT504000) AS AMT504000_SU, SUM(AMOUNT502000) AS AMT502000_SU
			FROM METRO_IT_MARGIN_DEPT
			WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) AND MERCH_GROUP_CODE = 'SU' 
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND STORE_FORMAT < 5
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY UPPER(STORE_FORMAT_DESC))B ON A.FORMAT = B.FORMAT
			)
		GROUP BY A.FORMAT ORDER BY 1
		});
	$sls1->execute();
						
		while(my $s = $sls1->fetchrow_hashref()){
								
			$worksheet->merge_range( $a, 4, $a, 6, $s->{FORMAT}, $desc );
						
			$worksheet->write($a,7, $s->{O_RETAIL_DS}+$s->{C_RETAIL_DS},$border1);
			$worksheet->write($a,8, $s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$border1);
				if (($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}) le 0){
					$worksheet->write($a,9, "",$subt); }
				else{
					$worksheet->write($a,9, ($s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}),$subt); }
						
			$worksheet->write($a,10, "",$border1);
			$worksheet->write($a,11, "",$subt);
			$worksheet->write($a,12, "",$border1);
			$worksheet->write($a,13, "",$subt);
			
			$worksheet->write($a,14, $s->{O_RETAIL_DS},$border1);
			$worksheet->write($a,15, $s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS},$border1);
				if ($s->{O_RETAIL_DS} le 0){
					$worksheet->write($a,16, "",$subt); }
				else{
					$worksheet->write($a,16, ($s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS})/$s->{O_RETAIL_DS},$subt); }
						
			$worksheet->write($a,17, "",$border1);
			$worksheet->write($a,18, "",$subt);
			$worksheet->write($a,19, "",$border1);
			$worksheet->write($a,20, "",$subt);			
			
			$worksheet->write($a,21, $s->{O_RETAIL_DS},$border1);
			$worksheet->write($a,22, $s->{O_MARGIN_DS},$border1);
				if ($s->{O_RETAIL_DS} le 0){
					$worksheet->write($a,23, "",$subt); }
				else{
					$worksheet->write($a,23, $s->{O_MARGIN_DS}/$s->{O_RETAIL_DS},$subt); }
					
			$worksheet->write($a,24, "",$border1);
			$worksheet->write($a,25, "",$subt);
			$worksheet->write($a,26, "",$border1);
			$worksheet->write($a,27, "",$subt);	
						
			$worksheet->write($a,28, $s->{AMT501000_DS},$border1);
			$worksheet->write($a,29, "",$border1);
			$worksheet->write($a,30, "",$border1);
			
			$worksheet->write($a,31, $s->{AMT503200_DS},$border1);
			$worksheet->write($a,32, "",$border1);
			$worksheet->write($a,33, "",$border1);
						
			$worksheet->write($a,34, $s->{AMT503250_DS},$border1);
			$worksheet->write($a,35, "",$border1);
			$worksheet->write($a,36, "",$border1);
					
			$worksheet->write($a,37, $s->{AMT503500_DS},$border1);
			$worksheet->write($a,38, "",$border1);
			$worksheet->write($a,39, "",$border1);
						
			$worksheet->write($a,40, $s->{AMT506000_DS},$border1);
			$worksheet->write($a,41, "",$border1);
			$worksheet->write($a,42, "",$border1);
						
			$worksheet->write($a,43, $s->{AMT503000_DS},$border1);
			$worksheet->write($a,44, "",$border1);
			$worksheet->write($a,45, "",$border1);				
								
			$worksheet->write($a,46, $s->{AMT507000_DS},$border1);
			$worksheet->write($a,47, "",$border1);
			$worksheet->write($a,48, "",$border1);
				
			$worksheet->write($a,49, $s->{AMT999998_DS},$border1);
			$worksheet->write($a,50, "",$border1);
			$worksheet->write($a,51, "",$border1);
			
			$worksheet->write($a,52, $s->{AMT504000_DS},$border1);
			$worksheet->write($a,53, "",$border1);
			$worksheet->write($a,54, "",$border1);
									
			$worksheet->write($a,55, $s->{AMT432000_DS},$border1);
			$worksheet->write($a,56, "",$border1);
			$worksheet->write($a,57, "",$border1);
						
			$worksheet->write($a,58, $s->{AMT433000_DS},$border1);
			$worksheet->write($a,59, "",$border1);
			$worksheet->write($a,60, "",$border1);
						
			$worksheet->write($a,61, $s->{AMT458490_DS},$border1);
			$worksheet->write($a,62, "",$border1);
			$worksheet->write($a,63, "",$border1);
				
			$worksheet->write($a,64, $s->{AMT505000_DS},$border1);
			$worksheet->write($a,65, "",$border1);
			$worksheet->write($a,66, "",$border1);
						
			$worksheet->write($a,67, $s->{C_RETAIL_DS},$border1);
			$worksheet->write($a,68, $s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$border1);
				if ($s->{C_RETAIL_DS} le 0){
					$worksheet->write($a,69, "",$subt); }
				else{
					$worksheet->write($a,69,($s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/$s->{C_RETAIL_DS},$subt); }
								
			$worksheet->write($a,70, "",$border1);
			$worksheet->write($a,71, "",$subt);
			$worksheet->write($a,72, "",$border1);
			$worksheet->write($a,73, "",$subt);	
				
			$worksheet->write($a,74, $s->{C_RETAIL_DS},$border1);
			$worksheet->write($a,75, $s->{C_MARGIN_DS},$border1);
				if ($s->{C_RETAIL_DS} le 0){
					$worksheet->write($a,76, "",$subt); }
				else{
					$worksheet->write($a,76, $s->{C_MARGIN_DS}/$s->{C_RETAIL_DS},$subt); }
							
			$worksheet->write($a,77, "",$border1);
			$worksheet->write($a,78, "",$subt);
			$worksheet->write($a,79, "",$border1);
			$worksheet->write($a,80, "",$subt);	
						
			$worksheet->write($a,81, $s->{AMT502000_DS},$border1);
			$worksheet->write($a,82, "",$border1);
			$worksheet->write($a,83, "",$border1);
					
			$worksheet->write($a,84, $s->{AMT434000_DS},$border1);
			$worksheet->write($a,85, "",$border1);
			$worksheet->write($a,86, "",$border1);
								
			$worksheet->write($a,87, $s->{AMT458550_DS},$border1);
			$worksheet->write($a,88, "",$border1);
			$worksheet->write($a,89, "",$border1);
						
			$worksheet->write($a,90, $s->{AMT460100_DS},$border1);
			$worksheet->write($a,91, "",$border1);
			$worksheet->write($a,92, "",$border1);
							
			$worksheet->write($a,93, $s->{AMT460200_DS},$border1);
			$worksheet->write($a,94, "",$border1);
			$worksheet->write($a,95, "",$border1);
								
			$worksheet->write($a,96, $s->{AMT460300_DS},$border1);
			$worksheet->write($a,97, "",$border1);
			$worksheet->write($a,98, "",$border1);	
			
			$worksheet->write($a,99, $s->{O_RETAIL_SU}+$s->{C_RETAIL_SU},$border1);
			$worksheet->write($a,100, $s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$border1);
				if (($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}) le 0){
					$worksheet->write($a,101, "",$subt); }
				else{
					$worksheet->write($a,101, ($s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}),$subt); }
						
			$worksheet->write($a,102, "",$border1);
			$worksheet->write($a,103, "",$subt);
			$worksheet->write($a,104, "",$border1);
			$worksheet->write($a,105, "",$subt);
			
			$worksheet->write($a,106, $s->{O_RETAIL_SU},$border1);
			$worksheet->write($a,107, $s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU},$border1);
				if ($s->{O_RETAIL_SU} le 0){
					$worksheet->write($a,108, "",$subt); }
				else{
					$worksheet->write($a,108, ($s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU})/$s->{O_RETAIL_SU},$subt); }
						
			$worksheet->write($a,109, "",$border1);
			$worksheet->write($a,110, "",$subt);
			$worksheet->write($a,111, "",$border1);
			$worksheet->write($a,112, "",$subt);			
			
			$worksheet->write($a,113, $s->{O_RETAIL_SU},$border1);
			$worksheet->write($a,114, $s->{O_MARGIN_SU},$border1);
				if ($s->{O_RETAIL_SU} le 0){
					$worksheet->write($a,115, "",$subt); }
				else{
					$worksheet->write($a,115, $s->{O_MARGIN_SU}/$s->{O_RETAIL_SU},$subt); }
					
			$worksheet->write($a,116, "",$border1);
			$worksheet->write($a,117, "",$subt);
			$worksheet->write($a,118, "",$border1);
			$worksheet->write($a,119, "",$subt);	
						
			$worksheet->write($a,120, $s->{AMT501000_SU},$border1);
			$worksheet->write($a,121, "",$border1);
			$worksheet->write($a,122, "",$border1);
			
			$worksheet->write($a,123, $s->{AMT503200_SU},$border1);
			$worksheet->write($a,124, "",$border1);
			$worksheet->write($a,125, "",$border1);
						
			$worksheet->write($a,126, $s->{AMT503250_SU},$border1);
			$worksheet->write($a,127, "",$border1);
			$worksheet->write($a,128, "",$border1);
					
			$worksheet->write($a,129, $s->{AMT503500_SU},$border1);
			$worksheet->write($a,130, "",$border1);
			$worksheet->write($a,131, "",$border1);
					
			$worksheet->write($a,132, $s->{AMT506000_SU},$border1);
			$worksheet->write($a,133, "",$border1);
			$worksheet->write($a,134, "",$border1);
						
			$worksheet->write($a,135, $s->{AMT503000_SU},$border1);
			$worksheet->write($a,136, "",$border1);
			$worksheet->write($a,137, "",$border1);				
								
			$worksheet->write($a,138, $s->{AMT507000_SU},$border1);
			$worksheet->write($a,139, "",$border1);
			$worksheet->write($a,140, "",$border1);
				
			$worksheet->write($a,141, $s->{AMT999998_SU},$border1);
			$worksheet->write($a,142, "",$border1);
			$worksheet->write($a,143, "",$border1);
			
			$worksheet->write($a,144, $s->{AMT504000_SU},$border1);
			$worksheet->write($a,145, "",$border1);
			$worksheet->write($a,146, "",$border1);
									
			$worksheet->write($a,147, $s->{AMT432000_SU},$border1);
			$worksheet->write($a,148, "",$border1);
			$worksheet->write($a,149, "",$border1);
						
			$worksheet->write($a,150, $s->{AMT433000_SU},$border1);
			$worksheet->write($a,151, "",$border1);
			$worksheet->write($a,152, "",$border1);
						
			$worksheet->write($a,153, $s->{AMT458490_SU},$border1);
			$worksheet->write($a,154, "",$border1);
			$worksheet->write($a,155, "",$border1);
				
			$worksheet->write($a,156, $s->{AMT505000_SU},$border1);
			$worksheet->write($a,157, "",$border1);
			$worksheet->write($a,158, "",$border1);
						
			$worksheet->write($a,159, $s->{C_RETAIL_SU},$border1);
			$worksheet->write($a,160, $s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$border1);
				if ($s->{C_RETAIL_SU} le 0){
					$worksheet->write($a,161, "",$subt); }
				else{
					$worksheet->write($a,161,($s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/$s->{C_RETAIL_SU},$subt); }
								
			$worksheet->write($a,162, "",$border1);
			$worksheet->write($a,163, "",$subt);
			$worksheet->write($a,164, "",$border1);
			$worksheet->write($a,165, "",$subt);	
				
			$worksheet->write($a,166, $s->{C_RETAIL_SU},$border1);
			$worksheet->write($a,167, $s->{C_MARGIN_SU},$border1);
				if ($s->{C_RETAIL_SU} le 0){
					$worksheet->write($a,168, "",$subt); }
				else{
					$worksheet->write($a,168, $s->{C_MARGIN_SU}/$s->{C_RETAIL_SU},$subt); }
							
			$worksheet->write($a,169, "",$border1);
			$worksheet->write($a,170, "",$subt);
			$worksheet->write($a,171, "",$border1);
			$worksheet->write($a,172, "",$subt);	
						
			$worksheet->write($a,173, $s->{AMT502000_SU},$border1);
			$worksheet->write($a,174, "",$border1);
			$worksheet->write($a,175, "",$border1);
					
			$worksheet->write($a,176, $s->{AMT434000_SU},$border1);
			$worksheet->write($a,177, "",$border1);
			$worksheet->write($a,178, "",$border1);
								
			$worksheet->write($a,179, $s->{AMT458550_SU},$border1);
			$worksheet->write($a,180, "",$border1);
			$worksheet->write($a,181, "",$border1);
						
			$worksheet->write($a,182, $s->{AMT460100_SU},$border1);
			$worksheet->write($a,183, "",$border1);
			$worksheet->write($a,184, "",$border1);
							
			$worksheet->write($a,185, $s->{AMT460200_SU},$border1);
			$worksheet->write($a,186, "",$border1);
			$worksheet->write($a,187, "",$border1);
								
			$worksheet->write($a,188, $s->{AMT460300_SU},$border1);
			$worksheet->write($a,189, "",$border1);
			$worksheet->write($a,190, "",$border1);	
												
			$a++;
			$counter++;
						
		}
	
	$worksheet->write($a,7, $s->{O_RETAIL_DS}+$s->{C_RETAIL_DS},$bodyNum);
	$worksheet->write($a,8, $s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$bodyNum);
		if (($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}) le 0){
			$worksheet->write($a,9, "",$bodyPct); }
		else{
			$worksheet->write($a,9, ($s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}),$bodyPct); }
						
	$worksheet->write($a,10, "",$bodyNum);
	$worksheet->write($a,11, "",$bodyPct);
	$worksheet->write($a,12, "",$bodyNum);
	$worksheet->write($a,13, "",$bodyPct);
	
	$worksheet->write($a,14, $s->{O_RETAIL_DS},$bodyNum);
	$worksheet->write($a,15, $s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS},$bodyNum);
		if ($s->{O_RETAIL_DS} le 0){
			$worksheet->write($a,16, "",$bodyPct); }
		else{
			$worksheet->write($a,16, ($s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS})/$s->{O_RETAIL_DS},$bodyPct); }
						
	$worksheet->write($a,17, "",$bodyNum);
	$worksheet->write($a,18, "",$bodyPct);
	$worksheet->write($a,19, "",$bodyNum);
	$worksheet->write($a,20, "",$bodyPct);			
	
	$worksheet->write($a,21, $s->{O_RETAIL_DS},$bodyNum);
	$worksheet->write($a,22, $s->{O_MARGIN_DS},$bodyNum);
		if ($s->{O_RETAIL_DS} le 0){
			$worksheet->write($a,23, "",$bodyPct); }
		else{
			$worksheet->write($a,23, $s->{O_MARGIN_DS}/$s->{O_RETAIL_DS},$bodyPct); }
			
	$worksheet->write($a,24, "",$bodyNum);
	$worksheet->write($a,25, "",$bodyPct);
	$worksheet->write($a,26, "",$bodyNum);
	$worksheet->write($a,27, "",$bodyPct);	
					
	$worksheet->write($a,28, $s->{AMT501000_DS},$bodyNum);
	$worksheet->write($a,29, "",$bodyNum);
	$worksheet->write($a,30, "",$bodyNum);
	
	$worksheet->write($a,31, $s->{AMT503200_DS},$bodyNum);
	$worksheet->write($a,32, "",$bodyNum);
	$worksheet->write($a,33, "",$bodyNum);
				
	$worksheet->write($a,34, $s->{AMT503250_DS},$bodyNum);
	$worksheet->write($a,35, "",$bodyNum);
	$worksheet->write($a,36, "",$bodyNum);
			
	$worksheet->write($a,37, $s->{AMT503500_DS},$bodyNum);
	$worksheet->write($a,38, "",$bodyNum);
	$worksheet->write($a,39, "",$bodyNum);
					
	$worksheet->write($a,40, $s->{AMT506000_DS},$bodyNum);
	$worksheet->write($a,41, "",$bodyNum);
	$worksheet->write($a,42, "",$bodyNum);
				
	$worksheet->write($a,43, $s->{AMT503000_DS},$bodyNum);
	$worksheet->write($a,44, "",$bodyNum);
	$worksheet->write($a,45, "",$bodyNum);				
							
	$worksheet->write($a,46, $s->{AMT507000_DS},$bodyNum);
	$worksheet->write($a,47, "",$bodyNum);
	$worksheet->write($a,48, "",$bodyNum);
				
	$worksheet->write($a,49, $s->{AMT999998_DS},$bodyNum);
	$worksheet->write($a,50, "",$bodyNum);
	$worksheet->write($a,51, "",$bodyNum);
			
	$worksheet->write($a,52, $s->{AMT504000_DS},$bodyNum);
	$worksheet->write($a,53, "",$bodyNum);
	$worksheet->write($a,54, "",$bodyNum);
									
	$worksheet->write($a,55, $s->{AMT432000_DS},$bodyNum);
	$worksheet->write($a,56, "",$bodyNum);
	$worksheet->write($a,57, "",$bodyNum);
						
	$worksheet->write($a,58, $s->{AMT433000_DS},$bodyNum);
	$worksheet->write($a,59, "",$bodyNum);
	$worksheet->write($a,60, "",$bodyNum);
						
	$worksheet->write($a,61, $s->{AMT458490_DS},$bodyNum);
	$worksheet->write($a,62, "",$bodyNum);
	$worksheet->write($a,63, "",$bodyNum);
			
	$worksheet->write($a,64, $s->{AMT505000_DS},$bodyNum);
	$worksheet->write($a,65, "",$bodyNum);
	$worksheet->write($a,66, "",$bodyNum);
				
	$worksheet->write($a,67, $s->{C_RETAIL_DS},$bodyNum);
	$worksheet->write($a,68, $s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$bodyNum);
		if ($s->{C_RETAIL_DS} le 0){
			$worksheet->write($a,69, "",$bodyPct); }
		else{
			$worksheet->write($a,69,($s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/$s->{C_RETAIL_DS},$bodyPct); }
								
	$worksheet->write($a,70, "",$bodyNum);
	$worksheet->write($a,71, "",$bodyPct);
	$worksheet->write($a,72, "",$bodyNum);
	$worksheet->write($a,73, "",$bodyPct);	
				
	$worksheet->write($a,74, $s->{C_RETAIL_DS},$bodyNum);
	$worksheet->write($a,75, $s->{C_MARGIN_DS},$bodyNum);
		if ($s->{C_RETAIL_DS} le 0){
			$worksheet->write($a,76, "",$bodyPct); }
		else{
			$worksheet->write($a,76, $s->{C_MARGIN_DS}/$s->{C_RETAIL_DS},$bodyPct); }
					
	$worksheet->write($a,77, "",$bodyNum);
	$worksheet->write($a,78, "",$bodyPct);
	$worksheet->write($a,79, "",$bodyNum);
	$worksheet->write($a,80, "",$bodyPct);	
						
	$worksheet->write($a,81, $s->{AMT502000_DS},$bodyNum);
	$worksheet->write($a,82, "",$bodyNum);
	$worksheet->write($a,83, "",$bodyNum);
			
	$worksheet->write($a,84, $s->{AMT434000_DS},$bodyNum);
	$worksheet->write($a,85, "",$bodyNum);
	$worksheet->write($a,86, "",$bodyNum);
						
	$worksheet->write($a,87, $s->{AMT458550_DS},$bodyNum);
	$worksheet->write($a,88, "",$bodyNum);
	$worksheet->write($a,89, "",$bodyNum);
				
	$worksheet->write($a,90, $s->{AMT460100_DS},$bodyNum);
	$worksheet->write($a,91, "",$bodyNum);
	$worksheet->write($a,92, "",$bodyNum);
					
	$worksheet->write($a,93, $s->{AMT460200_DS},$bodyNum);
	$worksheet->write($a,94, "",$bodyNum);
	$worksheet->write($a,95, "",$bodyNum);
						
	$worksheet->write($a,96, $s->{AMT460300_DS},$bodyNum);
	$worksheet->write($a,97, "",$bodyNum);
	$worksheet->write($a,98, "",$bodyNum);	
					
	$worksheet->write($a,99, $s->{O_RETAIL_SU}+$s->{C_RETAIL_SU},$bodyNum);
	$worksheet->write($a,100, $s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$bodyNum);
		if (($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}) le 0){
			$worksheet->write($a,101, "",$bodyPct); }
		else{
			$worksheet->write($a,101, ($s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}),$bodyPct); }
						
	$worksheet->write($a,102, "",$bodyNum);
	$worksheet->write($a,103, "",$bodyPct);
	$worksheet->write($a,104, "",$bodyNum);
	$worksheet->write($a,105, "",$bodyPct);
	
	$worksheet->write($a,106, $s->{O_RETAIL_SU},$bodyNum);
	$worksheet->write($a,107, $s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU},$bodyNum);
		if ($s->{O_RETAIL_SU} le 0){
			$worksheet->write($a,108, "",$bodyPct); }
		else{
			$worksheet->write($a,108, ($s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU})/$s->{O_RETAIL_SU},$bodyPct); }
						
	$worksheet->write($a,109, "",$bodyNum);
	$worksheet->write($a,110, "",$bodyPct);
	$worksheet->write($a,111, "",$bodyNum);
	$worksheet->write($a,112, "",$bodyPct);			
	
	$worksheet->write($a,113, $s->{O_RETAIL_SU},$bodyNum);
	$worksheet->write($a,114, $s->{O_MARGIN_SU},$bodyNum);
		if ($s->{O_RETAIL_SU} le 0){
			$worksheet->write($a,115, "",$bodyPct); }
		else{
			$worksheet->write($a,115, $s->{O_MARGIN_SU}/$s->{O_RETAIL_SU},$bodyPct); }
			
	$worksheet->write($a,116, "",$bodyNum);
	$worksheet->write($a,117, "",$bodyPct);
	$worksheet->write($a,118, "",$bodyNum);
	$worksheet->write($a,119, "",$bodyPct);	
				
	$worksheet->write($a,120, $s->{AMT501000_SU},$bodyNum);
	$worksheet->write($a,121, "",$bodyNum);
	$worksheet->write($a,122, "",$bodyNum);
		
	$worksheet->write($a,123, $s->{AMT503200_SU},$bodyNum);
	$worksheet->write($a,124, "",$bodyNum);
	$worksheet->write($a,125, "",$bodyNum);
						
	$worksheet->write($a,126, $s->{AMT503250_SU},$bodyNum);
	$worksheet->write($a,127, "",$bodyNum);
	$worksheet->write($a,128, "",$bodyNum);
					
	$worksheet->write($a,129, $s->{AMT503500_SU},$bodyNum);
	$worksheet->write($a,130, "",$bodyNum);
	$worksheet->write($a,131, "",$bodyNum);
					
	$worksheet->write($a,132, $s->{AMT506000_SU},$bodyNum);
	$worksheet->write($a,133, "",$bodyNum);
	$worksheet->write($a,134, "",$bodyNum);
						
	$worksheet->write($a,135, $s->{AMT503000_SU},$bodyNum);
	$worksheet->write($a,136, "",$bodyNum);
	$worksheet->write($a,137, "",$bodyNum);				
							
	$worksheet->write($a,138, $s->{AMT507000_SU},$bodyNum);
	$worksheet->write($a,139, "",$bodyNum);
	$worksheet->write($a,140, "",$bodyNum);
				
	$worksheet->write($a,141, $s->{AMT999998_SU},$bodyNum);
	$worksheet->write($a,142, "",$bodyNum);
	$worksheet->write($a,143, "",$bodyNum);
	
	$worksheet->write($a,144, $s->{AMT504000_SU},$bodyNum);
	$worksheet->write($a,145, "",$bodyNum);
	$worksheet->write($a,146, "",$bodyNum);
							
	$worksheet->write($a,147, $s->{AMT432000_SU},$bodyNum);
	$worksheet->write($a,148, "",$bodyNum);
	$worksheet->write($a,149, "",$bodyNum);
				
	$worksheet->write($a,150, $s->{AMT433000_SU},$bodyNum);
	$worksheet->write($a,151, "",$bodyNum);
	$worksheet->write($a,152, "",$bodyNum);
				
	$worksheet->write($a,153, $s->{AMT458490_SU},$bodyNum);
	$worksheet->write($a,154, "",$bodyNum);
	$worksheet->write($a,155, "",$bodyNum);
		
	$worksheet->write($a,156, $s->{AMT505000_SU},$bodyNum);
	$worksheet->write($a,157, "",$bodyNum);
	$worksheet->write($a,158, "",$bodyNum);
					
	$worksheet->write($a,159, $s->{C_RETAIL_SU},$bodyNum);
	$worksheet->write($a,160, $s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$bodyNum);
		if ($s->{C_RETAIL_SU} le 0){
			$worksheet->write($a,161, "",$bodyPct); }
		else{
			$worksheet->write($a,161,($s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/$s->{C_RETAIL_SU},$bodyPct); }
								
	$worksheet->write($a,162, "",$bodyNum);
	$worksheet->write($a,163, "",$bodyPct);
	$worksheet->write($a,164, "",$bodyNum);
	$worksheet->write($a,165, "",$bodyPct);	
				
	$worksheet->write($a,166, $s->{C_RETAIL_SU},$bodyNum);
	$worksheet->write($a,167, $s->{C_MARGIN_SU},$bodyNum);
		if ($s->{C_RETAIL_SU} le 0){
			$worksheet->write($a,168, "",$bodyPct); }
		else{
			$worksheet->write($a,168, $s->{C_MARGIN_SU}/$s->{C_RETAIL_SU},$bodyPct); }
							
	$worksheet->write($a,169, "",$bodyNum);
	$worksheet->write($a,170, "",$bodyPct);
	$worksheet->write($a,171, "",$bodyNum);
	$worksheet->write($a,172, "",$bodyPct);	
				
	$worksheet->write($a,173, $s->{AMT502000_SU},$bodyNum);
	$worksheet->write($a,174, "",$bodyNum);
	$worksheet->write($a,175, "",$bodyNum);
					
	$worksheet->write($a,176, $s->{AMT434000_SU},$bodyNum);
	$worksheet->write($a,177, "",$bodyNum);
	$worksheet->write($a,178, "",$bodyNum);
						
	$worksheet->write($a,179, $s->{AMT458550_SU},$bodyNum);
	$worksheet->write($a,180, "",$bodyNum);
	$worksheet->write($a,181, "",$bodyNum);
						
	$worksheet->write($a,182, $s->{AMT460100_SU},$bodyNum);
	$worksheet->write($a,183, "",$bodyNum);
	$worksheet->write($a,184, "",$bodyNum);
							
	$worksheet->write($a,185, $s->{AMT460200_SU},$bodyNum);
	$worksheet->write($a,186, "",$bodyNum);
	$worksheet->write($a,187, "",$bodyNum);
								
	$worksheet->write($a,188, $s->{AMT460300_SU},$bodyNum);
	$worksheet->write($a,189, "",$bodyNum);
	$worksheet->write($a,190, "",$bodyNum);
	
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
	SELECT SUM(O_COST_DS) AS O_COST_DS, SUM(O_RETAIL_DS) AS O_RETAIL_DS, SUM(O_MARGIN_DS) AS O_MARGIN_DS, SUM(C_COST_DS) AS C_COST_DS, SUM(C_RETAIL_DS) AS C_RETAIL_DS, SUM(C_MARGIN_DS) AS C_MARGIN_DS, SUM(AMT432000_DS) AS AMT432000_DS, SUM(AMT433000_DS) AS AMT433000_DS, SUM(AMT458490_DS) AS AMT458490_DS, SUM(AMT434000_DS) AS AMT434000_DS, SUM(AMT458550_DS) AS AMT458550_DS, SUM(AMT460100_DS) AS AMT460100_DS, SUM(AMT460200_DS) AS AMT460200_DS, SUM(AMT460300_DS) AS AMT460300_DS, SUM(AMT503200_DS) AS AMT503200_DS, SUM(AMT503250_DS) AS AMT503250_DS, SUM(AMT503500_DS) AS AMT503500_DS, SUM(AMT506000_DS) AS AMT506000_DS, SUM(AMT501000_DS) AS AMT501000_DS, SUM(AMT503000_DS) AS AMT503000_DS, SUM(AMT507000_DS) AS AMT507000_DS, SUM(AMT999998_DS) AS AMT999998_DS, SUM(AMT505000_DS) AS AMT505000_DS, SUM(AMT504000_DS) AS AMT504000_DS, SUM(AMT502000_DS) AS AMT502000_DS, SUM(O_COST_SU) AS O_COST_SU, SUM(O_RETAIL_SU) AS O_RETAIL_SU, SUM(O_MARGIN_SU) AS O_MARGIN_SU, SUM(C_COST_SU) AS C_COST_SU, SUM(C_RETAIL_SU) AS C_RETAIL_SU, SUM(C_MARGIN_SU) AS C_MARGIN_SU, SUM(AMT432000_SU) AS AMT432000_SU, SUM(AMT433000_SU) AS AMT433000_SU, SUM(AMT458490_SU) AS AMT458490_SU, SUM(AMT434000_SU) AS AMT434000_SU, SUM(AMT458550_SU) AS AMT458550_SU, SUM(AMT460100_SU) AS AMT460100_SU, SUM(AMT460200_SU) AS AMT460200_SU, SUM(AMT460300_SU) AS AMT460300_SU, SUM(AMT503200_SU) AS AMT503200_SU, SUM(AMT503250_SU) AS AMT503250_SU, SUM(AMT503500_SU) AS AMT503500_SU, SUM(AMT506000_SU) AS AMT506000_SU, SUM(AMT501000_SU) AS AMT501000_SU, SUM(AMT503000_SU) AS AMT503000_SU, SUM(AMT507000_SU) AS AMT507000_SU, SUM(AMT999998_SU) AS AMT999998_SU, SUM(AMT505000_SU) AS AMT505000_SU, SUM(AMT504000_SU) AS AMT504000_SU, SUM(AMT502000_SU) AS AMT502000_SU
	FROM
		(SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, SUM(OTR_COST) AS O_COST_DS, SUM(OTR_RETAIL) AS O_RETAIL_DS, SUM(OTR_MARGIN) AS O_MARGIN_DS, SUM(CON_COST) AS C_COST_DS, SUM(CON_RETAIL) AS C_RETAIL_DS, SUM(CON_MARGIN) AS C_MARGIN_DS, SUM(AMOUNT432000) AS AMT432000_DS, SUM(AMOUNT433000) AS AMT433000_DS, SUM(AMOUNT458490) AS AMT458490_DS, SUM(AMOUNT434000) AS AMT434000_DS, SUM(AMOUNT458550) AS AMT458550_DS, SUM(AMOUNT460100) AS AMT460100_DS, SUM(AMOUNT460200) AS AMT460200_DS, SUM(AMOUNT460300) AS AMT460300_DS, SUM(AMOUNT503200) AS AMT503200_DS, SUM(AMOUNT503250) AS AMT503250_DS, SUM(AMOUNT503500) AS AMT503500_DS, SUM(AMOUNT506000) AS AMT506000_DS, SUM(AMOUNT501000) AS AMT501000_DS, SUM(AMOUNT503000) AS AMT503000_DS, SUM(AMOUNT507000) AS AMT507000_DS, SUM(AMOUNT999998) AS AMT999998_DS, SUM(AMOUNT505000) AS AMT505000_DS, SUM(AMOUNT504000) AS AMT504000_DS, SUM(AMOUNT502000) AS AMT502000_DS, 0 AS O_COST_SU, 0 AS O_RETAIL_SU, 0 AS O_MARGIN_SU, 0 AS C_COST_SU, 0 AS C_RETAIL_SU, 0 AS C_MARGIN_SU, 0 AS AMT432000_SU, 0 AS AMT433000_SU, 0 AS AMT458490_SU, 0 AS AMT434000_SU, 0 AS AMT458550_SU, 0 AS AMT460100_SU, 0 AS AMT460200_SU, 0 AS AMT460300_SU, 0 AS AMT503200_SU, 0 AS AMT503250_SU, 0 AS AMT503500_SU, 0 AS AMT506000_SU, 0 AS AMT501000_SU, 0 AS AMT503000_SU, 0 AS AMT507000_SU, 0 AS AMT999998_SU, 0 AS AMT505000_SU, 0 AS AMT504000_SU, 0 AS AMT502000_SU
		FROM METRO_IT_MARGIN_DEPT
		WHERE MERCH_GROUP_CODE = 'DS' 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND UPPER(STORE_FORMAT_DESC) = '$store_format'
			AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
		GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_DESCRIPTION, MATURED_FLG
		UNION ALL
		SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 0 AS O_COST_DS, 0 AS O_RETAIL_DS, 0 AS O_MARGIN_DS, 0 AS C_COST_DS, 0 AS C_RETAIL_DS, 0 AS C_MARGIN_DS, 0 AS AMT432000_DS, 0 AS AMT433000_DS, 0 AS AMT458490_DS, 0 AS AMT434000_DS, 0 AS AMT458550_DS, 0 AS AMT460100_DS, 0 AS AMT460200_DS, 0 AS AMT460300_DS, 0 AS AMT503200_DS, 0 AS AMT503250_DS, 0 AS AMT503500_DS, 0 AS AMT506000_DS, 0 AS AMT501000_DS, 0 AS AMT503000_DS, 0 AS AMT507000_DS, 0 AS AMT999998_DS, 0 AS AMT505000_DS, 0 AS AMT504000_DS, 0 AS AMT502000_DS, SUM(OTR_COST) AS O_COST_SU, SUM(OTR_RETAIL) AS O_RETAIL_SU, SUM(OTR_MARGIN) AS O_MARGIN_SU, SUM(CON_COST) AS C_COST_SU, SUM(CON_RETAIL) AS C_RETAIL_SU, SUM(CON_MARGIN) AS C_MARGIN_SU, SUM(AMOUNT432000) AS AMT432000_SU, SUM(AMOUNT433000) AS AMT433000_SU, SUM(AMOUNT458490) AS AMT458490_SU, SUM(AMOUNT434000) AS AMT434000_SU, SUM(AMOUNT458550) AS AMT458550_SU, SUM(AMOUNT460100) AS AMT460100_SU, SUM(AMOUNT460200) AS AMT460200_SU, SUM(AMOUNT460300) AS AMT460300_SU, SUM(AMOUNT503200) AS AMT503200_SU, SUM(AMOUNT503250) AS AMT503250_SU, SUM(AMOUNT503500) AS AMT503500_SU, SUM(AMOUNT506000) AS AMT506000_SU, SUM(AMOUNT501000) AS AMT501000_SU, SUM(AMOUNT503000) AS AMT503000_SU, SUM(AMOUNT507000) AS AMT507000_SU, SUM(AMOUNT999998) AS AMT999998_SU, SUM(AMOUNT505000) AS AMT505000_SU, SUM(AMOUNT504000) AS AMT504000_SU, SUM(AMOUNT502000) AS AMT502000_SU
		FROM METRO_IT_MARGIN_DEPT
		WHERE MERCH_GROUP_CODE = 'SU' 
			AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND UPPER(STORE_FORMAT_DESC) = '$store_format'
			AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
		GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)	
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
		SELECT MATURED_FLG, CASE WHEN MATURED_FLG = 1 THEN 'COMP' ELSE 'NEW' END AS FLG_DESC, SUM(O_COST_DS) AS O_COST_DS, SUM(O_RETAIL_DS) AS O_RETAIL_DS, SUM(O_MARGIN_DS) AS O_MARGIN_DS, SUM(C_COST_DS) AS C_COST_DS, SUM(C_RETAIL_DS) AS C_RETAIL_DS, SUM(C_MARGIN_DS) AS C_MARGIN_DS, SUM(AMT432000_DS) AS AMT432000_DS, SUM(AMT433000_DS) AS AMT433000_DS, SUM(AMT458490_DS) AS AMT458490_DS, SUM(AMT434000_DS) AS AMT434000_DS, SUM(AMT458550_DS) AS AMT458550_DS, SUM(AMT460100_DS) AS AMT460100_DS, SUM(AMT460200_DS) AS AMT460200_DS, SUM(AMT460300_DS) AS AMT460300_DS, SUM(AMT503200_DS) AS AMT503200_DS, SUM(AMT503250_DS) AS AMT503250_DS, SUM(AMT503500_DS) AS AMT503500_DS, SUM(AMT506000_DS) AS AMT506000_DS, SUM(AMT501000_DS) AS AMT501000_DS, SUM(AMT503000_DS) AS AMT503000_DS, SUM(AMT507000_DS) AS AMT507000_DS, SUM(AMT999998_DS) AS AMT999998_DS, SUM(AMT505000_DS) AS AMT505000_DS, SUM(AMT504000_DS) AS AMT504000_DS, SUM(AMT502000_DS) AS AMT502000_DS, SUM(O_COST_SU) AS O_COST_SU, SUM(O_RETAIL_SU) AS O_RETAIL_SU, SUM(O_MARGIN_SU) AS O_MARGIN_SU, SUM(C_COST_SU) AS C_COST_SU, SUM(C_RETAIL_SU) AS C_RETAIL_SU, SUM(C_MARGIN_SU) AS C_MARGIN_SU, SUM(AMT432000_SU) AS AMT432000_SU, SUM(AMT433000_SU) AS AMT433000_SU, SUM(AMT458490_SU) AS AMT458490_SU, SUM(AMT434000_SU) AS AMT434000_SU, SUM(AMT458550_SU) AS AMT458550_SU, SUM(AMT460100_SU) AS AMT460100_SU, SUM(AMT460200_SU) AS AMT460200_SU, SUM(AMT460300_SU) AS AMT460300_SU, SUM(AMT503200_SU) AS AMT503200_SU, SUM(AMT503250_SU) AS AMT503250_SU, SUM(AMT503500_SU) AS AMT503500_SU, SUM(AMT506000_SU) AS AMT506000_SU, SUM(AMT501000_SU) AS AMT501000_SU, SUM(AMT503000_SU) AS AMT503000_SU, SUM(AMT507000_SU) AS AMT507000_SU, SUM(AMT999998_SU) AS AMT999998_SU, SUM(AMT505000_SU) AS AMT505000_SU, SUM(AMT504000_SU) AS AMT504000_SU, SUM(AMT502000_SU) AS AMT502000_SU
		FROM
			(SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, SUM(OTR_COST) AS O_COST_DS, SUM(OTR_RETAIL) AS O_RETAIL_DS, SUM(OTR_MARGIN) AS O_MARGIN_DS, SUM(CON_COST) AS C_COST_DS, SUM(CON_RETAIL) AS C_RETAIL_DS, SUM(CON_MARGIN) AS C_MARGIN_DS, SUM(AMOUNT432000) AS AMT432000_DS, SUM(AMOUNT433000) AS AMT433000_DS, SUM(AMOUNT458490) AS AMT458490_DS, SUM(AMOUNT434000) AS AMT434000_DS, SUM(AMOUNT458550) AS AMT458550_DS, SUM(AMOUNT460100) AS AMT460100_DS, SUM(AMOUNT460200) AS AMT460200_DS, SUM(AMOUNT460300) AS AMT460300_DS, SUM(AMOUNT503200) AS AMT503200_DS, SUM(AMOUNT503250) AS AMT503250_DS, SUM(AMOUNT503500) AS AMT503500_DS, SUM(AMOUNT506000) AS AMT506000_DS, SUM(AMOUNT501000) AS AMT501000_DS, SUM(AMOUNT503000) AS AMT503000_DS, SUM(AMOUNT507000) AS AMT507000_DS, SUM(AMOUNT999998) AS AMT999998_DS, SUM(AMOUNT505000) AS AMT505000_DS, SUM(AMOUNT504000) AS AMT504000_DS, SUM(AMOUNT502000) AS AMT502000_DS, 0 AS O_COST_SU, 0 AS O_RETAIL_SU, 0 AS O_MARGIN_SU, 0 AS C_COST_SU, 0 AS C_RETAIL_SU, 0 AS C_MARGIN_SU, 0 AS AMT432000_SU, 0 AS AMT433000_SU, 0 AS AMT458490_SU, 0 AS AMT434000_SU, 0 AS AMT458550_SU, 0 AS AMT460100_SU, 0 AS AMT460200_SU, 0 AS AMT460300_SU, 0 AS AMT503200_SU, 0 AS AMT503250_SU, 0 AS AMT503500_SU, 0 AS AMT506000_SU, 0 AS AMT501000_SU, 0 AS AMT503000_SU, 0 AS AMT507000_SU, 0 AS AMT999998_SU, 0 AS AMT505000_SU, 0 AS AMT504000_SU, 0 AS AMT502000_SU
			FROM METRO_IT_MARGIN_DEPT
			WHERE MERCH_GROUP_CODE = 'DS' 
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND UPPER(STORE_FORMAT_DESC) = '$store_format'
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_DESCRIPTION, MATURED_FLG
			UNION ALL
			SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 0 AS O_COST_DS, 0 AS O_RETAIL_DS, 0 AS O_MARGIN_DS, 0 AS C_COST_DS, 0 AS C_RETAIL_DS, 0 AS C_MARGIN_DS, 0 AS AMT432000_DS, 0 AS AMT433000_DS, 0 AS AMT458490_DS, 0 AS AMT434000_DS, 0 AS AMT458550_DS, 0 AS AMT460100_DS, 0 AS AMT460200_DS, 0 AS AMT460300_DS, 0 AS AMT503200_DS, 0 AS AMT503250_DS, 0 AS AMT503500_DS, 0 AS AMT506000_DS, 0 AS AMT501000_DS, 0 AS AMT503000_DS, 0 AS AMT507000_DS, 0 AS AMT999998_DS, 0 AS AMT505000_DS, 0 AS AMT504000_DS, 0 AS AMT502000_DS, SUM(OTR_COST) AS O_COST_SU, SUM(OTR_RETAIL) AS O_RETAIL_SU, SUM(OTR_MARGIN) AS O_MARGIN_SU, SUM(CON_COST) AS C_COST_SU, SUM(CON_RETAIL) AS C_RETAIL_SU, SUM(CON_MARGIN) AS C_MARGIN_SU, SUM(AMOUNT432000) AS AMT432000_SU, SUM(AMOUNT433000) AS AMT433000_SU, SUM(AMOUNT458490) AS AMT458490_SU, SUM(AMOUNT434000) AS AMT434000_SU, SUM(AMOUNT458550) AS AMT458550_SU, SUM(AMOUNT460100) AS AMT460100_SU, SUM(AMOUNT460200) AS AMT460200_SU, SUM(AMOUNT460300) AS AMT460300_SU, SUM(AMOUNT503200) AS AMT503200_SU, SUM(AMOUNT503250) AS AMT503250_SU, SUM(AMOUNT503500) AS AMT503500_SU, SUM(AMOUNT506000) AS AMT506000_SU, SUM(AMOUNT501000) AS AMT501000_SU, SUM(AMOUNT503000) AS AMT503000_SU, SUM(AMOUNT507000) AS AMT507000_SU, SUM(AMOUNT999998) AS AMT999998_SU, SUM(AMOUNT505000) AS AMT505000_SU, SUM(AMOUNT504000) AS AMT504000_SU, SUM(AMOUNT502000) AS AMT502000_SU
			FROM METRO_IT_MARGIN_DEPT
			WHERE MERCH_GROUP_CODE = 'SU' 
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND UPPER(STORE_FORMAT_DESC) = '$store_format'
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)
		GROUP BY MATURED_FLG, CASE WHEN MATURED_FLG = 1 THEN 'COMP' ELSE 'NEW' END
		ORDER BY MATURED_FLG DESC
	});
	$sls1->execute();
		
		$format_counter = $a;
		while(my $s = $sls1->fetchrow_hashref()){
		$flg = $s->{MATURED_FLG};
		$flg_desc = $s->{FLG_DESC};

		$sls2 = $dbh->prepare (qq{
			SELECT STORE_CODE, STORE_DESCRIPTION, SUM(O_COST_DS) AS O_COST_DS, SUM(O_RETAIL_DS) AS O_RETAIL_DS, SUM(O_MARGIN_DS) AS O_MARGIN_DS, SUM(C_COST_DS) AS C_COST_DS, SUM(C_RETAIL_DS) AS C_RETAIL_DS, SUM(C_MARGIN_DS) AS C_MARGIN_DS, SUM(AMT432000_DS) AS AMT432000_DS, SUM(AMT433000_DS) AS AMT433000_DS, SUM(AMT458490_DS) AS AMT458490_DS, SUM(AMT434000_DS) AS AMT434000_DS, SUM(AMT458550_DS) AS AMT458550_DS, SUM(AMT460100_DS) AS AMT460100_DS, SUM(AMT460200_DS) AS AMT460200_DS, SUM(AMT460300_DS) AS AMT460300_DS, SUM(AMT503200_DS) AS AMT503200_DS, SUM(AMT503250_DS) AS AMT503250_DS, SUM(AMT503500_DS) AS AMT503500_DS, SUM(AMT506000_DS) AS AMT506000_DS, SUM(AMT501000_DS) AS AMT501000_DS, SUM(AMT503000_DS) AS AMT503000_DS, SUM(AMT507000_DS) AS AMT507000_DS, SUM(AMT999998_DS) AS AMT999998_DS, SUM(AMT505000_DS) AS AMT505000_DS, SUM(AMT504000_DS) AS AMT504000_DS, SUM(AMT502000_DS) AS AMT502000_DS, SUM(O_COST_SU) AS O_COST_SU, SUM(O_RETAIL_SU) AS O_RETAIL_SU, SUM(O_MARGIN_SU) AS O_MARGIN_SU, SUM(C_COST_SU) AS C_COST_SU, SUM(C_RETAIL_SU) AS C_RETAIL_SU, SUM(C_MARGIN_SU) AS C_MARGIN_SU, SUM(AMT432000_SU) AS AMT432000_SU, SUM(AMT433000_SU) AS AMT433000_SU, SUM(AMT458490_SU) AS AMT458490_SU, SUM(AMT434000_SU) AS AMT434000_SU, SUM(AMT458550_SU) AS AMT458550_SU, SUM(AMT460100_SU) AS AMT460100_SU, SUM(AMT460200_SU) AS AMT460200_SU, SUM(AMT460300_SU) AS AMT460300_SU, SUM(AMT503200_SU) AS AMT503200_SU, SUM(AMT503250_SU) AS AMT503250_SU, SUM(AMT503500_SU) AS AMT503500_SU, SUM(AMT506000_SU) AS AMT506000_SU, SUM(AMT501000_SU) AS AMT501000_SU, SUM(AMT503000_SU) AS AMT503000_SU, SUM(AMT507000_SU) AS AMT507000_SU, SUM(AMT999998_SU) AS AMT999998_SU, SUM(AMT505000_SU) AS AMT505000_SU, SUM(AMT504000_SU) AS AMT504000_SU, SUM(AMT502000_SU) AS AMT502000_SU
			FROM
				(SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, SUM(OTR_COST) AS O_COST_DS, SUM(OTR_RETAIL) AS O_RETAIL_DS, SUM(OTR_MARGIN) AS O_MARGIN_DS, SUM(CON_COST) AS C_COST_DS, SUM(CON_RETAIL) AS C_RETAIL_DS, SUM(CON_MARGIN) AS C_MARGIN_DS, SUM(AMOUNT432000) AS AMT432000_DS, SUM(AMOUNT433000) AS AMT433000_DS, SUM(AMOUNT458490) AS AMT458490_DS, SUM(AMOUNT434000) AS AMT434000_DS, SUM(AMOUNT458550) AS AMT458550_DS, SUM(AMOUNT460100) AS AMT460100_DS, SUM(AMOUNT460200) AS AMT460200_DS, SUM(AMOUNT460300) AS AMT460300_DS, SUM(AMOUNT503200) AS AMT503200_DS, SUM(AMOUNT503250) AS AMT503250_DS, SUM(AMOUNT503500) AS AMT503500_DS, SUM(AMOUNT506000) AS AMT506000_DS, SUM(AMOUNT501000) AS AMT501000_DS, SUM(AMOUNT503000) AS AMT503000_DS, SUM(AMOUNT507000) AS AMT507000_DS, SUM(AMOUNT999998) AS AMT999998_DS, SUM(AMOUNT505000) AS AMT505000_DS, SUM(AMOUNT504000) AS AMT504000_DS, SUM(AMOUNT502000) AS AMT502000_DS, 0 AS O_COST_SU, 0 AS O_RETAIL_SU, 0 AS O_MARGIN_SU, 0 AS C_COST_SU, 0 AS C_RETAIL_SU, 0 AS C_MARGIN_SU, 0 AS AMT432000_SU, 0 AS AMT433000_SU, 0 AS AMT458490_SU, 0 AS AMT434000_SU, 0 AS AMT458550_SU, 0 AS AMT460100_SU, 0 AS AMT460200_SU, 0 AS AMT460300_SU, 0 AS AMT503200_SU, 0 AS AMT503250_SU, 0 AS AMT503500_SU, 0 AS AMT506000_SU, 0 AS AMT501000_SU, 0 AS AMT503000_SU, 0 AS AMT507000_SU, 0 AS AMT999998_SU, 0 AS AMT505000_SU, 0 AS AMT504000_SU, 0 AS AMT502000_SU
				FROM METRO_IT_MARGIN_DEPT
				WHERE MERCH_GROUP_CODE = 'DS' 
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND UPPER(STORE_FORMAT_DESC) = '$store_format'
					AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') ) AND MATURED_FLG = '$flg'
				GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_DESCRIPTION, MATURED_FLG
				UNION ALL
				SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_DESCRIPTION, MATURED_FLG, 0 AS O_COST_DS, 0 AS O_RETAIL_DS, 0 AS O_MARGIN_DS, 0 AS C_COST_DS, 0 AS C_RETAIL_DS, 0 AS C_MARGIN_DS, 0 AS AMT432000_DS, 0 AS AMT433000_DS, 0 AS AMT458490_DS, 0 AS AMT434000_DS, 0 AS AMT458550_DS, 0 AS AMT460100_DS, 0 AS AMT460200_DS, 0 AS AMT460300_DS, 0 AS AMT503200_DS, 0 AS AMT503250_DS, 0 AS AMT503500_DS, 0 AS AMT506000_DS, 0 AS AMT501000_DS, 0 AS AMT503000_DS, 0 AS AMT507000_DS, 0 AS AMT999998_DS, 0 AS AMT505000_DS, 0 AS AMT504000_DS, 0 AS AMT502000_DS, SUM(OTR_COST) AS O_COST_SU, SUM(OTR_RETAIL) AS O_RETAIL_SU, SUM(OTR_MARGIN) AS O_MARGIN_SU, SUM(CON_COST) AS C_COST_SU, SUM(CON_RETAIL) AS C_RETAIL_SU, SUM(CON_MARGIN) AS C_MARGIN_SU, SUM(AMOUNT432000) AS AMT432000_SU, SUM(AMOUNT433000) AS AMT433000_SU, SUM(AMOUNT458490) AS AMT458490_SU, SUM(AMOUNT434000) AS AMT434000_SU, SUM(AMOUNT458550) AS AMT458550_SU, SUM(AMOUNT460100) AS AMT460100_SU, SUM(AMOUNT460200) AS AMT460200_SU, SUM(AMOUNT460300) AS AMT460300_SU, SUM(AMOUNT503200) AS AMT503200_SU, SUM(AMOUNT503250) AS AMT503250_SU, SUM(AMOUNT503500) AS AMT503500_SU, SUM(AMOUNT506000) AS AMT506000_SU, SUM(AMOUNT501000) AS AMT501000_SU, SUM(AMOUNT503000) AS AMT503000_SU, SUM(AMOUNT507000) AS AMT507000_SU, SUM(AMOUNT999998) AS AMT999998_SU, SUM(AMOUNT505000) AS AMT505000_SU, SUM(AMOUNT504000) AS AMT504000_SU, SUM(AMOUNT502000) AS AMT502000_SU
				FROM METRO_IT_MARGIN_DEPT
				WHERE MERCH_GROUP_CODE = 'SU' 
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON'))) AND UPPER(STORE_FORMAT_DESC) = '$store_format'
					AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') ) AND MATURED_FLG = '$flg'
				GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_DESCRIPTION, MATURED_FLG)
			GROUP BY STORE_CODE, STORE_DESCRIPTION
			ORDER BY 1		
		});
		$sls2->execute();
			
			while(my $s = $sls2->fetchrow_hashref()){
									
			$worksheet->write( $a, 5, $s->{STORE_CODE}, $desc );
			$worksheet->write( $a, 6, $s->{STORE_DESCRIPTION}, $desc );
			$worksheet->write($a,7, $s->{O_RETAIL_DS}+$s->{C_RETAIL_DS},$border1);
			$worksheet->write($a,8, $s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$border1);
				if (($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}) le 0){
					$worksheet->write($a,9, "",$subt); }
				else{
					$worksheet->write($a,9, ($s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}),$subt); }
						
			$worksheet->write($a,10, "",$border1);
			$worksheet->write($a,11, "",$subt);
			$worksheet->write($a,12, "",$border1);
			$worksheet->write($a,13, "",$subt);
			
			$worksheet->write($a,14, $s->{O_RETAIL_DS},$border1);
			$worksheet->write($a,15, $s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS},$border1);
				if ($s->{O_RETAIL_DS} le 0){
					$worksheet->write($a,16, "",$subt); }
				else{
					$worksheet->write($a,16, ($s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS})/$s->{O_RETAIL_DS},$subt); }
						
			$worksheet->write($a,17, "",$border1);
			$worksheet->write($a,18, "",$subt);
			$worksheet->write($a,19, "",$border1);
			$worksheet->write($a,20, "",$subt);			
			
			$worksheet->write($a,21, $s->{O_RETAIL_DS},$border1);
			$worksheet->write($a,22, $s->{O_MARGIN_DS},$border1);
				if ($s->{O_RETAIL_DS} le 0){
					$worksheet->write($a,23, "",$subt); }
				else{
					$worksheet->write($a,23, $s->{O_MARGIN_DS}/$s->{O_RETAIL_DS},$subt); }
					
			$worksheet->write($a,24, "",$border1);
			$worksheet->write($a,25, "",$subt);
			$worksheet->write($a,26, "",$border1);
			$worksheet->write($a,27, "",$subt);	
						
			$worksheet->write($a,28, $s->{AMT501000_DS},$border1);
			$worksheet->write($a,29, "",$border1);
			$worksheet->write($a,30, "",$border1);
			
			$worksheet->write($a,31, $s->{AMT503200_DS},$border1);
			$worksheet->write($a,32, "",$border1);
			$worksheet->write($a,33, "",$border1);
						
			$worksheet->write($a,34, $s->{AMT503250_DS},$border1);
			$worksheet->write($a,35, "",$border1);
			$worksheet->write($a,36, "",$border1);
					
			$worksheet->write($a,37, $s->{AMT503500_DS},$border1);
			$worksheet->write($a,38, "",$border1);
			$worksheet->write($a,39, "",$border1);
						
			$worksheet->write($a,40, $s->{AMT506000_DS},$border1);
			$worksheet->write($a,41, "",$border1);
			$worksheet->write($a,42, "",$border1);
						
			$worksheet->write($a,43, $s->{AMT503000_DS},$border1);
			$worksheet->write($a,44, "",$border1);
			$worksheet->write($a,45, "",$border1);				
								
			$worksheet->write($a,46, $s->{AMT507000_DS},$border1);
			$worksheet->write($a,47, "",$border1);
			$worksheet->write($a,48, "",$border1);
				
			$worksheet->write($a,49, $s->{AMT999998_DS},$border1);
			$worksheet->write($a,50, "",$border1);
			$worksheet->write($a,51, "",$border1);
			
			$worksheet->write($a,52, $s->{AMT504000_DS},$border1);
			$worksheet->write($a,53, "",$border1);
			$worksheet->write($a,54, "",$border1);
									
			$worksheet->write($a,55, $s->{AMT432000_DS},$border1);
			$worksheet->write($a,56, "",$border1);
			$worksheet->write($a,57, "",$border1);
						
			$worksheet->write($a,58, $s->{AMT433000_DS},$border1);
			$worksheet->write($a,59, "",$border1);
			$worksheet->write($a,60, "",$border1);
						
			$worksheet->write($a,61, $s->{AMT458490_DS},$border1);
			$worksheet->write($a,62, "",$border1);
			$worksheet->write($a,63, "",$border1);
				
			$worksheet->write($a,64, $s->{AMT505000_DS},$border1);
			$worksheet->write($a,65, "",$border1);
			$worksheet->write($a,66, "",$border1);
						
			$worksheet->write($a,67, $s->{C_RETAIL_DS},$border1);
			$worksheet->write($a,68, $s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$border1);
				if ($s->{C_RETAIL_DS} le 0){
					$worksheet->write($a,69, "",$subt); }
				else{
					$worksheet->write($a,69,($s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/$s->{C_RETAIL_DS},$subt); }
								
			$worksheet->write($a,70, "",$border1);
			$worksheet->write($a,71, "",$subt);
			$worksheet->write($a,72, "",$border1);
			$worksheet->write($a,73, "",$subt);	
				
			$worksheet->write($a,74, $s->{C_RETAIL_DS},$border1);
			$worksheet->write($a,75, $s->{C_MARGIN_DS},$border1);
				if ($s->{C_RETAIL_DS} le 0){
					$worksheet->write($a,76, "",$subt); }
				else{
					$worksheet->write($a,76, $s->{C_MARGIN_DS}/$s->{C_RETAIL_DS},$subt); }
							
			$worksheet->write($a,77, "",$border1);
			$worksheet->write($a,78, "",$subt);
			$worksheet->write($a,79, "",$border1);
			$worksheet->write($a,80, "",$subt);	
						
			$worksheet->write($a,81, $s->{AMT502000_DS},$border1);
			$worksheet->write($a,82, "",$border1);
			$worksheet->write($a,83, "",$border1);
					
			$worksheet->write($a,84, $s->{AMT434000_DS},$border1);
			$worksheet->write($a,85, "",$border1);
			$worksheet->write($a,86, "",$border1);
								
			$worksheet->write($a,87, $s->{AMT458550_DS},$border1);
			$worksheet->write($a,88, "",$border1);
			$worksheet->write($a,89, "",$border1);
						
			$worksheet->write($a,90, $s->{AMT460100_DS},$border1);
			$worksheet->write($a,91, "",$border1);
			$worksheet->write($a,92, "",$border1);
							
			$worksheet->write($a,93, $s->{AMT460200_DS},$border1);
			$worksheet->write($a,94, "",$border1);
			$worksheet->write($a,95, "",$border1);
								
			$worksheet->write($a,96, $s->{AMT460300_DS},$border1);
			$worksheet->write($a,97, "",$border1);
			$worksheet->write($a,98, "",$border1);	
						
			$worksheet->write($a,99, $s->{O_RETAIL_SU}+$s->{C_RETAIL_SU},$border1);
			$worksheet->write($a,100, $s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$border1);
				if (($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}) le 0){
					$worksheet->write($a,101, "",$subt); }
				else{
					$worksheet->write($a,101, ($s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}),$subt); }
						
			$worksheet->write($a,102, "",$border1);
			$worksheet->write($a,103, "",$subt);
			$worksheet->write($a,104, "",$border1);
			$worksheet->write($a,105, "",$subt);
			
			$worksheet->write($a,106, $s->{O_RETAIL_SU},$border1);
			$worksheet->write($a,107, $s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU},$border1);
				if ($s->{O_RETAIL_SU} le 0){
					$worksheet->write($a,108, "",$subt); }
				else{
					$worksheet->write($a,108, ($s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU})/$s->{O_RETAIL_SU},$subt); }
						
			$worksheet->write($a,109, "",$border1);
			$worksheet->write($a,110, "",$subt);
			$worksheet->write($a,111, "",$border1);
			$worksheet->write($a,112, "",$subt);			
			
			$worksheet->write($a,113, $s->{O_RETAIL_SU},$border1);
			$worksheet->write($a,114, $s->{O_MARGIN_SU},$border1);
				if ($s->{O_RETAIL_SU} le 0){
					$worksheet->write($a,115, "",$subt); }
				else{
					$worksheet->write($a,115, $s->{O_MARGIN_SU}/$s->{O_RETAIL_SU},$subt); }
					
			$worksheet->write($a,116, "",$border1);
			$worksheet->write($a,117, "",$subt);
			$worksheet->write($a,118, "",$border1);
			$worksheet->write($a,119, "",$subt);	
						
			$worksheet->write($a,120, $s->{AMT501000_SU},$border1);
			$worksheet->write($a,121, "",$border1);
			$worksheet->write($a,122, "",$border1);
			
			$worksheet->write($a,123, $s->{AMT503200_SU},$border1);
			$worksheet->write($a,124, "",$border1);
			$worksheet->write($a,125, "",$border1);
						
			$worksheet->write($a,126, $s->{AMT503250_SU},$border1);
			$worksheet->write($a,127, "",$border1);
			$worksheet->write($a,128, "",$border1);
					
			$worksheet->write($a,129, $s->{AMT503500_SU},$border1);
			$worksheet->write($a,130, "",$border1);
			$worksheet->write($a,131, "",$border1);
					
			$worksheet->write($a,132, $s->{AMT506000_SU},$border1);
			$worksheet->write($a,133, "",$border1);
			$worksheet->write($a,134, "",$border1);
						
			$worksheet->write($a,135, $s->{AMT503000_SU},$border1);
			$worksheet->write($a,136, "",$border1);
			$worksheet->write($a,137, "",$border1);				
								
			$worksheet->write($a,138, $s->{AMT507000_SU},$border1);
			$worksheet->write($a,139, "",$border1);
			$worksheet->write($a,140, "",$border1);
				
			$worksheet->write($a,141, $s->{AMT999998_SU},$border1);
			$worksheet->write($a,142, "",$border1);
			$worksheet->write($a,143, "",$border1);
			
			$worksheet->write($a,144, $s->{AMT504000_SU},$border1);
			$worksheet->write($a,145, "",$border1);
			$worksheet->write($a,146, "",$border1);
									
			$worksheet->write($a,147, $s->{AMT432000_SU},$border1);
			$worksheet->write($a,148, "",$border1);
			$worksheet->write($a,149, "",$border1);
						
			$worksheet->write($a,150, $s->{AMT433000_SU},$border1);
			$worksheet->write($a,151, "",$border1);
			$worksheet->write($a,152, "",$border1);
						
			$worksheet->write($a,153, $s->{AMT458490_SU},$border1);
			$worksheet->write($a,154, "",$border1);
			$worksheet->write($a,155, "",$border1);
				
			$worksheet->write($a,156, $s->{AMT505000_SU},$border1);
			$worksheet->write($a,157, "",$border1);
			$worksheet->write($a,158, "",$border1);
						
			$worksheet->write($a,159, $s->{C_RETAIL_SU},$border1);
			$worksheet->write($a,160, $s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$border1);
				if ($s->{C_RETAIL_SU} le 0){
					$worksheet->write($a,161, "",$subt); }
				else{
					$worksheet->write($a,161,($s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/$s->{C_RETAIL_SU},$subt); }
								
			$worksheet->write($a,162, "",$border1);
			$worksheet->write($a,163, "",$subt);
			$worksheet->write($a,164, "",$border1);
			$worksheet->write($a,165, "",$subt);	
				
			$worksheet->write($a,166, $s->{C_RETAIL_SU},$border1);
			$worksheet->write($a,167, $s->{C_MARGIN_SU},$border1);
				if ($s->{C_RETAIL_SU} le 0){
					$worksheet->write($a,168, "",$subt); }
				else{
					$worksheet->write($a,168, $s->{C_MARGIN_SU}/$s->{C_RETAIL_SU},$subt); }
							
			$worksheet->write($a,169, "",$border1);
			$worksheet->write($a,170, "",$subt);
			$worksheet->write($a,171, "",$border1);
			$worksheet->write($a,172, "",$subt);	
						
			$worksheet->write($a,173, $s->{AMT502000_SU},$border1);
			$worksheet->write($a,174, "",$border1);
			$worksheet->write($a,175, "",$border1);
					
			$worksheet->write($a,176, $s->{AMT434000_SU},$border1);
			$worksheet->write($a,177, "",$border1);
			$worksheet->write($a,178, "",$border1);
								
			$worksheet->write($a,179, $s->{AMT458550_SU},$border1);
			$worksheet->write($a,180, "",$border1);
			$worksheet->write($a,181, "",$border1);
						
			$worksheet->write($a,182, $s->{AMT460100_SU},$border1);
			$worksheet->write($a,183, "",$border1);
			$worksheet->write($a,184, "",$border1);
							
			$worksheet->write($a,185, $s->{AMT460200_SU},$border1);
			$worksheet->write($a,186, "",$border1);
			$worksheet->write($a,187, "",$border1);
								
			$worksheet->write($a,188, $s->{AMT460300_SU},$border1);
			$worksheet->write($a,189, "",$border1);
			$worksheet->write($a,190, "",$border1);
													
			$a++;
			$counter++;
						
		}
		
		$worksheet->write($a,7, $s->{O_RETAIL_DS}+$s->{C_RETAIL_DS},$bodyNum);
		$worksheet->write($a,8, $s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$bodyNum);
			if (($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}) le 0){
				$worksheet->write($a,9, "",$bodyPct); }
			else{
				$worksheet->write($a,9, ($s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}),$bodyPct); }
							
		$worksheet->write($a,10, "",$bodyNum);
		$worksheet->write($a,11, "",$bodyPct);
		$worksheet->write($a,12, "",$bodyNum);
		$worksheet->write($a,13, "",$bodyPct);
		
		$worksheet->write($a,14, $s->{O_RETAIL_DS},$bodyNum);
		$worksheet->write($a,15, $s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS},$bodyNum);
			if ($s->{O_RETAIL_DS} le 0){
				$worksheet->write($a,16, "",$bodyPct); }
			else{
				$worksheet->write($a,16, ($s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS})/$s->{O_RETAIL_DS},$bodyPct); }
							
		$worksheet->write($a,17, "",$bodyNum);
		$worksheet->write($a,18, "",$bodyPct);
		$worksheet->write($a,19, "",$bodyNum);
		$worksheet->write($a,20, "",$bodyPct);			
		
		$worksheet->write($a,21, $s->{O_RETAIL_DS},$bodyNum);
		$worksheet->write($a,22, $s->{O_MARGIN_DS},$bodyNum);
			if ($s->{O_RETAIL_DS} le 0){
				$worksheet->write($a,23, "",$bodyPct); }
			else{
				$worksheet->write($a,23, $s->{O_MARGIN_DS}/$s->{O_RETAIL_DS},$bodyPct); }
				
		$worksheet->write($a,24, "",$bodyNum);
		$worksheet->write($a,25, "",$bodyPct);
		$worksheet->write($a,26, "",$bodyNum);
		$worksheet->write($a,27, "",$bodyPct);	
						
		$worksheet->write($a,28, $s->{AMT501000_DS},$bodyNum);
		$worksheet->write($a,29, "",$bodyNum);
		$worksheet->write($a,30, "",$bodyNum);
		
		$worksheet->write($a,31, $s->{AMT503200_DS},$bodyNum);
		$worksheet->write($a,32, "",$bodyNum);
		$worksheet->write($a,33, "",$bodyNum);
					
		$worksheet->write($a,34, $s->{AMT503250_DS},$bodyNum);
		$worksheet->write($a,35, "",$bodyNum);
		$worksheet->write($a,36, "",$bodyNum);
				
		$worksheet->write($a,37, $s->{AMT503500_DS},$bodyNum);
		$worksheet->write($a,38, "",$bodyNum);
		$worksheet->write($a,39, "",$bodyNum);
						
		$worksheet->write($a,40, $s->{AMT506000_DS},$bodyNum);
		$worksheet->write($a,41, "",$bodyNum);
		$worksheet->write($a,42, "",$bodyNum);
					
		$worksheet->write($a,43, $s->{AMT503000_DS},$bodyNum);
		$worksheet->write($a,44, "",$bodyNum);
		$worksheet->write($a,45, "",$bodyNum);				
								
		$worksheet->write($a,46, $s->{AMT507000_DS},$bodyNum);
		$worksheet->write($a,47, "",$bodyNum);
		$worksheet->write($a,48, "",$bodyNum);
					
		$worksheet->write($a,49, $s->{AMT999998_DS},$bodyNum);
		$worksheet->write($a,50, "",$bodyNum);
		$worksheet->write($a,51, "",$bodyNum);
				
		$worksheet->write($a,52, $s->{AMT504000_DS},$bodyNum);
		$worksheet->write($a,53, "",$bodyNum);
		$worksheet->write($a,54, "",$bodyNum);
										
		$worksheet->write($a,55, $s->{AMT432000_DS},$bodyNum);
		$worksheet->write($a,56, "",$bodyNum);
		$worksheet->write($a,57, "",$bodyNum);
							
		$worksheet->write($a,58, $s->{AMT433000_DS},$bodyNum);
		$worksheet->write($a,59, "",$bodyNum);
		$worksheet->write($a,60, "",$bodyNum);
							
		$worksheet->write($a,61, $s->{AMT458490_DS},$bodyNum);
		$worksheet->write($a,62, "",$bodyNum);
		$worksheet->write($a,63, "",$bodyNum);
				
		$worksheet->write($a,64, $s->{AMT505000_DS},$bodyNum);
		$worksheet->write($a,65, "",$bodyNum);
		$worksheet->write($a,66, "",$bodyNum);
					
		$worksheet->write($a,67, $s->{C_RETAIL_DS},$bodyNum);
		$worksheet->write($a,68, $s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$bodyNum);
			if ($s->{C_RETAIL_DS} le 0){
				$worksheet->write($a,69, "",$bodyPct); }
			else{
				$worksheet->write($a,69,($s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/$s->{C_RETAIL_DS},$bodyPct); }
									
		$worksheet->write($a,70, "",$bodyNum);
		$worksheet->write($a,71, "",$bodyPct);
		$worksheet->write($a,72, "",$bodyNum);
		$worksheet->write($a,73, "",$bodyPct);	
					
		$worksheet->write($a,74, $s->{C_RETAIL_DS},$bodyNum);
		$worksheet->write($a,75, $s->{C_MARGIN_DS},$bodyNum);
			if ($s->{C_RETAIL_DS} le 0){
				$worksheet->write($a,76, "",$bodyPct); }
			else{
				$worksheet->write($a,76, $s->{C_MARGIN_DS}/$s->{C_RETAIL_DS},$bodyPct); }
						
		$worksheet->write($a,77, "",$bodyNum);
		$worksheet->write($a,78, "",$bodyPct);
		$worksheet->write($a,79, "",$bodyNum);
		$worksheet->write($a,80, "",$bodyPct);	
							
		$worksheet->write($a,81, $s->{AMT502000_DS},$bodyNum);
		$worksheet->write($a,82, "",$bodyNum);
		$worksheet->write($a,83, "",$bodyNum);
				
		$worksheet->write($a,84, $s->{AMT434000_DS},$bodyNum);
		$worksheet->write($a,85, "",$bodyNum);
		$worksheet->write($a,86, "",$bodyNum);
							
		$worksheet->write($a,87, $s->{AMT458550_DS},$bodyNum);
		$worksheet->write($a,88, "",$bodyNum);
		$worksheet->write($a,89, "",$bodyNum);
					
		$worksheet->write($a,90, $s->{AMT460100_DS},$bodyNum);
		$worksheet->write($a,91, "",$bodyNum);
		$worksheet->write($a,92, "",$bodyNum);
						
		$worksheet->write($a,93, $s->{AMT460200_DS},$bodyNum);
		$worksheet->write($a,94, "",$bodyNum);
		$worksheet->write($a,95, "",$bodyNum);
							
		$worksheet->write($a,96, $s->{AMT460300_DS},$bodyNum);
		$worksheet->write($a,97, "",$bodyNum);
		$worksheet->write($a,98, "",$bodyNum);	
						
		$worksheet->write($a,99, $s->{O_RETAIL_SU}+$s->{C_RETAIL_SU},$bodyNum);
		$worksheet->write($a,100, $s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$bodyNum);
			if (($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}) le 0){
				$worksheet->write($a,101, "",$bodyPct); }
			else{
				$worksheet->write($a,101, ($s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}),$bodyPct); }
							
		$worksheet->write($a,102, "",$bodyNum);
		$worksheet->write($a,103, "",$bodyPct);
		$worksheet->write($a,104, "",$bodyNum);
		$worksheet->write($a,105, "",$bodyPct);
		
		$worksheet->write($a,106, $s->{O_RETAIL_SU},$bodyNum);
		$worksheet->write($a,107, $s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU},$bodyNum);
			if ($s->{O_RETAIL_SU} le 0){
				$worksheet->write($a,108, "",$bodyPct); }
			else{
				$worksheet->write($a,108, ($s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU})/$s->{O_RETAIL_SU},$bodyPct); }
							
		$worksheet->write($a,109, "",$bodyNum);
		$worksheet->write($a,110, "",$bodyPct);
		$worksheet->write($a,111, "",$bodyNum);
		$worksheet->write($a,112, "",$bodyPct);			
		
		$worksheet->write($a,113, $s->{O_RETAIL_SU},$bodyNum);
		$worksheet->write($a,114, $s->{O_MARGIN_SU},$bodyNum);
			if ($s->{O_RETAIL_SU} le 0){
				$worksheet->write($a,115, "",$bodyPct); }
			else{
				$worksheet->write($a,115, $s->{O_MARGIN_SU}/$s->{O_RETAIL_SU},$bodyPct); }
				
		$worksheet->write($a,116, "",$bodyNum);
		$worksheet->write($a,117, "",$bodyPct);
		$worksheet->write($a,118, "",$bodyNum);
		$worksheet->write($a,119, "",$bodyPct);	
					
		$worksheet->write($a,120, $s->{AMT501000_SU},$bodyNum);
		$worksheet->write($a,121, "",$bodyNum);
		$worksheet->write($a,122, "",$bodyNum);
			
		$worksheet->write($a,123, $s->{AMT503200_SU},$bodyNum);
		$worksheet->write($a,124, "",$bodyNum);
		$worksheet->write($a,125, "",$bodyNum);
							
		$worksheet->write($a,126, $s->{AMT503250_SU},$bodyNum);
		$worksheet->write($a,127, "",$bodyNum);
		$worksheet->write($a,128, "",$bodyNum);
						
		$worksheet->write($a,129, $s->{AMT503500_SU},$bodyNum);
		$worksheet->write($a,130, "",$bodyNum);
		$worksheet->write($a,131, "",$bodyNum);
						
		$worksheet->write($a,132, $s->{AMT506000_SU},$bodyNum);
		$worksheet->write($a,133, "",$bodyNum);
		$worksheet->write($a,134, "",$bodyNum);
							
		$worksheet->write($a,135, $s->{AMT503000_SU},$bodyNum);
		$worksheet->write($a,136, "",$bodyNum);
		$worksheet->write($a,137, "",$bodyNum);				
								
		$worksheet->write($a,138, $s->{AMT507000_SU},$bodyNum);
		$worksheet->write($a,139, "",$bodyNum);
		$worksheet->write($a,140, "",$bodyNum);
					
		$worksheet->write($a,141, $s->{AMT999998_SU},$bodyNum);
		$worksheet->write($a,142, "",$bodyNum);
		$worksheet->write($a,143, "",$bodyNum);
	
		$worksheet->write($a,144, $s->{AMT504000_SU},$bodyNum);
		$worksheet->write($a,145, "",$bodyNum);
		$worksheet->write($a,146, "",$bodyNum);
								
		$worksheet->write($a,147, $s->{AMT432000_SU},$bodyNum);
		$worksheet->write($a,148, "",$bodyNum);
		$worksheet->write($a,149, "",$bodyNum);
					
		$worksheet->write($a,150, $s->{AMT433000_SU},$bodyNum);
		$worksheet->write($a,151, "",$bodyNum);
		$worksheet->write($a,152, "",$bodyNum);
					
		$worksheet->write($a,153, $s->{AMT458490_SU},$bodyNum);
		$worksheet->write($a,154, "",$bodyNum);
		$worksheet->write($a,155, "",$bodyNum);
			
		$worksheet->write($a,156, $s->{AMT505000_SU},$bodyNum);
		$worksheet->write($a,157, "",$bodyNum);
		$worksheet->write($a,158, "",$bodyNum);
						
		$worksheet->write($a,159, $s->{C_RETAIL_SU},$bodyNum);
		$worksheet->write($a,160, $s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$bodyNum);
			if ($s->{C_RETAIL_SU} le 0){
				$worksheet->write($a,161, "",$bodyPct); }
			else{
				$worksheet->write($a,161,($s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/$s->{C_RETAIL_SU},$bodyPct); }
									
		$worksheet->write($a,162, "",$bodyNum);
		$worksheet->write($a,163, "",$bodyPct);
		$worksheet->write($a,164, "",$bodyNum);
		$worksheet->write($a,165, "",$bodyPct);	
					
		$worksheet->write($a,166, $s->{C_RETAIL_SU},$bodyNum);
		$worksheet->write($a,167, $s->{C_MARGIN_SU},$bodyNum);
			if ($s->{C_RETAIL_SU} le 0){
				$worksheet->write($a,168, "",$bodyPct); }
			else{
				$worksheet->write($a,168, $s->{C_MARGIN_SU}/$s->{C_RETAIL_SU},$bodyPct); }
								
		$worksheet->write($a,169, "",$bodyNum);
		$worksheet->write($a,170, "",$bodyPct);
		$worksheet->write($a,171, "",$bodyNum);
		$worksheet->write($a,172, "",$bodyPct);	
					
		$worksheet->write($a,173, $s->{AMT502000_SU},$bodyNum);
		$worksheet->write($a,174, "",$bodyNum);
		$worksheet->write($a,175, "",$bodyNum);
						
		$worksheet->write($a,176, $s->{AMT434000_SU},$bodyNum);
		$worksheet->write($a,177, "",$bodyNum);
		$worksheet->write($a,178, "",$bodyNum);
							
		$worksheet->write($a,179, $s->{AMT458550_SU},$bodyNum);
		$worksheet->write($a,180, "",$bodyNum);
		$worksheet->write($a,181, "",$bodyNum);
							
		$worksheet->write($a,182, $s->{AMT460100_SU},$bodyNum);
		$worksheet->write($a,183, "",$bodyNum);
		$worksheet->write($a,184, "",$bodyNum);
								
		$worksheet->write($a,185, $s->{AMT460200_SU},$bodyNum);
		$worksheet->write($a,186, "",$bodyNum);
		$worksheet->write($a,187, "",$bodyNum);
									
		$worksheet->write($a,188, $s->{AMT460300_SU},$bodyNum);
		$worksheet->write($a,189, "",$bodyNum);
		$worksheet->write($a,190, "",$bodyNum);
		
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $flg_desc, $border2 );
		
		$a++;
		$counter = 0;
	}
	
	$worksheet->write($a,7, $s->{O_RETAIL_DS}+$s->{C_RETAIL_DS},$bodyNum);
	$worksheet->write($a,8, $s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$bodyNum);
		if (($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}) le 0){
			$worksheet->write($a,9, "",$bodyPct); }
		else{
			$worksheet->write($a,9, ($s->{O_MARGIN_DS}+$s->{C_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/($s->{O_RETAIL_DS}+$s->{C_RETAIL_DS}),$bodyPct); }
						
	$worksheet->write($a,10, "",$bodyNum);
	$worksheet->write($a,11, "",$bodyPct);
	$worksheet->write($a,12, "",$bodyNum);
	$worksheet->write($a,13, "",$bodyPct);
	
	$worksheet->write($a,14, $s->{O_RETAIL_DS},$bodyNum);
	$worksheet->write($a,15, $s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS},$bodyNum);
		if ($s->{O_RETAIL_DS} le 0){
			$worksheet->write($a,16, "",$bodyPct); }
		else{
			$worksheet->write($a,16, ($s->{O_MARGIN_DS}+$s->{AMT501000_DS}+$s->{AMT432000_DS}+$s->{AMT433000_DS}+$s->{AMT458490_DS}+$s->{AMT503200_DS}+$s->{AMT503250_DS}+$s->{AMT503500_DS}+$s->{AMT506000_DS}+$s->{AMT503000_DS}+$s->{AMT507000_DS}+$s->{AMT999998_DS}+$s->{AMT505000_DS}+$s->{AMT504000_DS})/$s->{O_RETAIL_DS},$bodyPct); }
						
	$worksheet->write($a,17, "",$bodyNum);
	$worksheet->write($a,18, "",$bodyPct);
	$worksheet->write($a,19, "",$bodyNum);
	$worksheet->write($a,20, "",$bodyPct);			
	
	$worksheet->write($a,21, $s->{O_RETAIL_DS},$bodyNum);
	$worksheet->write($a,22, $s->{O_MARGIN_DS},$bodyNum);
		if ($s->{O_RETAIL_DS} le 0){
			$worksheet->write($a,23, "",$bodyPct); }
		else{
			$worksheet->write($a,23, $s->{O_MARGIN_DS}/$s->{O_RETAIL_DS},$bodyPct); }
			
	$worksheet->write($a,24, "",$bodyNum);
	$worksheet->write($a,25, "",$bodyPct);
	$worksheet->write($a,26, "",$bodyNum);
	$worksheet->write($a,27, "",$bodyPct);	
					
	$worksheet->write($a,28, $s->{AMT501000_DS},$bodyNum);
	$worksheet->write($a,29, "",$bodyNum);
	$worksheet->write($a,30, "",$bodyNum);
	
	$worksheet->write($a,31, $s->{AMT503200_DS},$bodyNum);
	$worksheet->write($a,32, "",$bodyNum);
	$worksheet->write($a,33, "",$bodyNum);
				
	$worksheet->write($a,34, $s->{AMT503250_DS},$bodyNum);
	$worksheet->write($a,35, "",$bodyNum);
	$worksheet->write($a,36, "",$bodyNum);
			
	$worksheet->write($a,37, $s->{AMT503500_DS},$bodyNum);
	$worksheet->write($a,38, "",$bodyNum);
	$worksheet->write($a,39, "",$bodyNum);
					
	$worksheet->write($a,40, $s->{AMT506000_DS},$bodyNum);
	$worksheet->write($a,41, "",$bodyNum);
	$worksheet->write($a,42, "",$bodyNum);
				
	$worksheet->write($a,43, $s->{AMT503000_DS},$bodyNum);
	$worksheet->write($a,44, "",$bodyNum);
	$worksheet->write($a,45, "",$bodyNum);				
							
	$worksheet->write($a,46, $s->{AMT507000_DS},$bodyNum);
	$worksheet->write($a,47, "",$bodyNum);
	$worksheet->write($a,48, "",$bodyNum);
				
	$worksheet->write($a,49, $s->{AMT999998_DS},$bodyNum);
	$worksheet->write($a,50, "",$bodyNum);
	$worksheet->write($a,51, "",$bodyNum);
			
	$worksheet->write($a,52, $s->{AMT504000_DS},$bodyNum);
	$worksheet->write($a,53, "",$bodyNum);
	$worksheet->write($a,54, "",$bodyNum);
									
	$worksheet->write($a,55, $s->{AMT432000_DS},$bodyNum);
	$worksheet->write($a,56, "",$bodyNum);
	$worksheet->write($a,57, "",$bodyNum);
						
	$worksheet->write($a,58, $s->{AMT433000_DS},$bodyNum);
	$worksheet->write($a,59, "",$bodyNum);
	$worksheet->write($a,60, "",$bodyNum);
						
	$worksheet->write($a,61, $s->{AMT458490_DS},$bodyNum);
	$worksheet->write($a,62, "",$bodyNum);
	$worksheet->write($a,63, "",$bodyNum);
			
	$worksheet->write($a,64, $s->{AMT505000_DS},$bodyNum);
	$worksheet->write($a,65, "",$bodyNum);
	$worksheet->write($a,66, "",$bodyNum);
				
	$worksheet->write($a,67, $s->{C_RETAIL_DS},$bodyNum);
	$worksheet->write($a,68, $s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS},$bodyNum);
		if ($s->{C_RETAIL_DS} le 0){
			$worksheet->write($a,69, "",$bodyPct); }
		else{
			$worksheet->write($a,69,($s->{C_MARGIN_DS}+$s->{AMT434000_DS}+$s->{AMT458550_DS}+$s->{AMT460100_DS}+$s->{AMT460200_DS}+$s->{AMT460300_DS}+$s->{AMT502000_DS})/$s->{C_RETAIL_DS},$bodyPct); }
								
	$worksheet->write($a,70, "",$bodyNum);
	$worksheet->write($a,71, "",$bodyPct);
	$worksheet->write($a,72, "",$bodyNum);
	$worksheet->write($a,73, "",$bodyPct);	
				
	$worksheet->write($a,74, $s->{C_RETAIL_DS},$bodyNum);
	$worksheet->write($a,75, $s->{C_MARGIN_DS},$bodyNum);
		if ($s->{C_RETAIL_DS} le 0){
			$worksheet->write($a,76, "",$bodyPct); }
		else{
			$worksheet->write($a,76, $s->{C_MARGIN_DS}/$s->{C_RETAIL_DS},$bodyPct); }
					
	$worksheet->write($a,77, "",$bodyNum);
	$worksheet->write($a,78, "",$bodyPct);
	$worksheet->write($a,79, "",$bodyNum);
	$worksheet->write($a,80, "",$bodyPct);	
						
	$worksheet->write($a,81, $s->{AMT502000_DS},$bodyNum);
	$worksheet->write($a,82, "",$bodyNum);
	$worksheet->write($a,83, "",$bodyNum);
			
	$worksheet->write($a,84, $s->{AMT434000_DS},$bodyNum);
	$worksheet->write($a,85, "",$bodyNum);
	$worksheet->write($a,86, "",$bodyNum);
						
	$worksheet->write($a,87, $s->{AMT458550_DS},$bodyNum);
	$worksheet->write($a,88, "",$bodyNum);
	$worksheet->write($a,89, "",$bodyNum);
				
	$worksheet->write($a,90, $s->{AMT460100_DS},$bodyNum);
	$worksheet->write($a,91, "",$bodyNum);
	$worksheet->write($a,92, "",$bodyNum);
					
	$worksheet->write($a,93, $s->{AMT460200_DS},$bodyNum);
	$worksheet->write($a,94, "",$bodyNum);
	$worksheet->write($a,95, "",$bodyNum);
						
	$worksheet->write($a,96, $s->{AMT460300_DS},$bodyNum);
	$worksheet->write($a,97, "",$bodyNum);
	$worksheet->write($a,98, "",$bodyNum);	
					
	$worksheet->write($a,99, $s->{O_RETAIL_SU}+$s->{C_RETAIL_SU},$bodyNum);
	$worksheet->write($a,100, $s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$bodyNum);
		if (($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}) le 0){
			$worksheet->write($a,101, "",$bodyPct); }
		else{
			$worksheet->write($a,101, ($s->{O_MARGIN_SU}+$s->{C_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/($s->{O_RETAIL_SU}+$s->{C_RETAIL_SU}),$bodyPct); }
						
	$worksheet->write($a,102, "",$bodyNum);
	$worksheet->write($a,103, "",$bodyPct);
	$worksheet->write($a,104, "",$bodyNum);
	$worksheet->write($a,105, "",$bodyPct);
	
	$worksheet->write($a,106, $s->{O_RETAIL_SU},$bodyNum);
	$worksheet->write($a,107, $s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU},$bodyNum);
		if ($s->{O_RETAIL_SU} le 0){
			$worksheet->write($a,108, "",$bodyPct); }
		else{
			$worksheet->write($a,108, ($s->{O_MARGIN_SU}+$s->{AMT501000_SU}+$s->{AMT432000_SU}+$s->{AMT433000_SU}+$s->{AMT458490_SU}+$s->{AMT503200_SU}+$s->{AMT503250_SU}+$s->{AMT503500_SU}+$s->{AMT506000_SU}+$s->{AMT503000_SU}+$s->{AMT507000_SU}+$s->{AMT999998_SU}+$s->{AMT505000_SU}+$s->{AMT504000_SU})/$s->{O_RETAIL_SU},$bodyPct); }
						
	$worksheet->write($a,109, "",$bodyNum);
	$worksheet->write($a,110, "",$bodyPct);
	$worksheet->write($a,111, "",$bodyNum);
	$worksheet->write($a,112, "",$bodyPct);			
	
	$worksheet->write($a,113, $s->{O_RETAIL_SU},$bodyNum);
	$worksheet->write($a,114, $s->{O_MARGIN_SU},$bodyNum);
		if ($s->{O_RETAIL_SU} le 0){
			$worksheet->write($a,115, "",$bodyPct); }
		else{
			$worksheet->write($a,115, $s->{O_MARGIN_SU}/$s->{O_RETAIL_SU},$bodyPct); }
			
	$worksheet->write($a,116, "",$bodyNum);
	$worksheet->write($a,117, "",$bodyPct);
	$worksheet->write($a,118, "",$bodyNum);
	$worksheet->write($a,119, "",$bodyPct);	
				
	$worksheet->write($a,120, $s->{AMT501000_SU},$bodyNum);
	$worksheet->write($a,121, "",$bodyNum);
	$worksheet->write($a,122, "",$bodyNum);
		
	$worksheet->write($a,123, $s->{AMT503200_SU},$bodyNum);
	$worksheet->write($a,124, "",$bodyNum);
	$worksheet->write($a,125, "",$bodyNum);
						
	$worksheet->write($a,126, $s->{AMT503250_SU},$bodyNum);
	$worksheet->write($a,127, "",$bodyNum);
	$worksheet->write($a,128, "",$bodyNum);
					
	$worksheet->write($a,129, $s->{AMT503500_SU},$bodyNum);
	$worksheet->write($a,130, "",$bodyNum);
	$worksheet->write($a,131, "",$bodyNum);
					
	$worksheet->write($a,132, $s->{AMT506000_SU},$bodyNum);
	$worksheet->write($a,133, "",$bodyNum);
	$worksheet->write($a,134, "",$bodyNum);
						
	$worksheet->write($a,135, $s->{AMT503000_SU},$bodyNum);
	$worksheet->write($a,136, "",$bodyNum);
	$worksheet->write($a,137, "",$bodyNum);				
							
	$worksheet->write($a,138, $s->{AMT507000_SU},$bodyNum);
	$worksheet->write($a,139, "",$bodyNum);
	$worksheet->write($a,140, "",$bodyNum);
				
	$worksheet->write($a,141, $s->{AMT999998_SU},$bodyNum);
	$worksheet->write($a,142, "",$bodyNum);
	$worksheet->write($a,143, "",$bodyNum);

	$worksheet->write($a,144, $s->{AMT504000_SU},$bodyNum);
	$worksheet->write($a,145, "",$bodyNum);
	$worksheet->write($a,146, "",$bodyNum);
							
	$worksheet->write($a,147, $s->{AMT432000_SU},$bodyNum);
	$worksheet->write($a,148, "",$bodyNum);
	$worksheet->write($a,149, "",$bodyNum);
				
	$worksheet->write($a,150, $s->{AMT433000_SU},$bodyNum);
	$worksheet->write($a,151, "",$bodyNum);
	$worksheet->write($a,152, "",$bodyNum);
				
	$worksheet->write($a,153, $s->{AMT458490_SU},$bodyNum);
	$worksheet->write($a,154, "",$bodyNum);
	$worksheet->write($a,155, "",$bodyNum);
		
	$worksheet->write($a,156, $s->{AMT505000_SU},$bodyNum);
	$worksheet->write($a,157, "",$bodyNum);
	$worksheet->write($a,158, "",$bodyNum);
					
	$worksheet->write($a,159, $s->{C_RETAIL_SU},$bodyNum);
	$worksheet->write($a,160, $s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU},$bodyNum);
		if ($s->{C_RETAIL_SU} le 0){
			$worksheet->write($a,161, "",$bodyPct); }
		else{
			$worksheet->write($a,161,($s->{C_MARGIN_SU}+$s->{AMT434000_SU}+$s->{AMT458550_SU}+$s->{AMT460100_SU}+$s->{AMT460200_SU}+$s->{AMT460300_SU}+$s->{AMT502000_SU})/$s->{C_RETAIL_SU},$bodyPct); }
								
	$worksheet->write($a,162, "",$bodyNum);
	$worksheet->write($a,163, "",$bodyPct);
	$worksheet->write($a,164, "",$bodyNum);
	$worksheet->write($a,165, "",$bodyPct);	
				
	$worksheet->write($a,166, $s->{C_RETAIL_SU},$bodyNum);
	$worksheet->write($a,167, $s->{C_MARGIN_SU},$bodyNum);
		if ($s->{C_RETAIL_SU} le 0){
			$worksheet->write($a,168, "",$bodyPct); }
		else{
			$worksheet->write($a,168, $s->{C_MARGIN_SU}/$s->{C_RETAIL_SU},$bodyPct); }
							
	$worksheet->write($a,169, "",$bodyNum);
	$worksheet->write($a,170, "",$bodyPct);
	$worksheet->write($a,171, "",$bodyNum);
	$worksheet->write($a,172, "",$bodyPct);	
				
	$worksheet->write($a,173, $s->{AMT502000_SU},$bodyNum);
	$worksheet->write($a,174, "",$bodyNum);
	$worksheet->write($a,175, "",$bodyNum);
					
	$worksheet->write($a,176, $s->{AMT434000_SU},$bodyNum);
	$worksheet->write($a,177, "",$bodyNum);
	$worksheet->write($a,178, "",$bodyNum);
						
	$worksheet->write($a,179, $s->{AMT458550_SU},$bodyNum);
	$worksheet->write($a,180, "",$bodyNum);
	$worksheet->write($a,181, "",$bodyNum);
						
	$worksheet->write($a,182, $s->{AMT460100_SU},$bodyNum);
	$worksheet->write($a,183, "",$bodyNum);
	$worksheet->write($a,184, "",$bodyNum);
							
	$worksheet->write($a,185, $s->{AMT460200_SU},$bodyNum);
	$worksheet->write($a,186, "",$bodyNum);
	$worksheet->write($a,187, "",$bodyNum);
								
	$worksheet->write($a,188, $s->{AMT460300_SU},$bodyNum);
	$worksheet->write($a,189, "",$bodyNum);
	$worksheet->write($a,190, "",$bodyNum);
	
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
			SELECT SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
			FROM METRO_IT_MARGIN_DEPT
			WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
				AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
			});								 
$sls->execute();

	while(my $s = $sls->fetchrow_hashref()){
				
	$sls1 = $dbh->prepare (qq{
				SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
				FROM METRO_IT_MARGIN_DEPT
				WHERE ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
				GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC
				ORDER BY 1
				});								 
	$sls1->execute();

		while(my $s = $sls1->fetchrow_hashref()){
			$merch_group_code = $s->{MERCH_GROUP_CODE};
			$merch_group_desc = $s->{MERCH_GROUP_DESC}; 
			
			$sls2 = $dbh->prepare (qq{
					SELECT GROUP_CODE, GROUP_DESC, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
					FROM METRO_IT_MARGIN_DEPT
					WHERE MERCH_GROUP_CODE = '$merch_group_code' 
						AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
						AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
					GROUP BY GROUP_CODE, GROUP_DESC
					ORDER BY 1
					});	
			$sls2->execute();
			
			$mgc_counter = $a;
			while(my $s = $sls2->fetchrow_hashref()){
				$group_code = $s->{GROUP_CODE};
				$group_desc = $s->{GROUP_DESC};
						
				$sls3 = $dbh->prepare (qq{
						SELECT DIVISION, DIVISION_DESC, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
						FROM METRO_IT_MARGIN_DEPT
						WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
							AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
						GROUP BY DIVISION, DIVISION_DESC
						ORDER BY 1
					});
				$sls3->execute();
				
				$grp_counter = $a;
				while(my $s = $sls3->fetchrow_hashref()){
					$division = $s->{DIVISION};
					$division_desc = $s->{DIVISION_DESC};
					
					$sls4 = $dbh->prepare (qq{	 
							SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
							 FROM METRO_IT_MARGIN_DEPT
							 WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2'))
								AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
							 GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC
							 ORDER BY 1
						});
					$sls4->execute();
					
					while(my $s = $sls4->fetchrow_hashref()){
						
						$worksheet->write($a,5, $s->{DEPARTMENT_CODE},$desc);
						$worksheet->write($a,6, $s->{DEPARTMENT_DESC},$desc);
						
						$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$border1);
						$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$border1);
							if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
								$worksheet->write($a,9, "",$subt); }
							else{
								$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$subt); }
						
						$worksheet->write($a,10, "",$border1);
						$worksheet->write($a,11, "",$subt);
						$worksheet->write($a,12, "",$border1);
						$worksheet->write($a,13, "",$subt);
						
						$worksheet->write($a,14, $s->{O_RETAIL},$border1);
						$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$border1);
							if ($s->{O_RETAIL} le 0){
								$worksheet->write($a,16, "",$subt); }
							else{
								$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$subt); }
						
						$worksheet->write($a,17, "",$border1);
						$worksheet->write($a,18, "",$subt);
						$worksheet->write($a,19, "",$border1);
						$worksheet->write($a,20, "",$subt);			
						
						$worksheet->write($a,21, $s->{O_RETAIL},$border1);
						$worksheet->write($a,22, $s->{O_MARGIN},$border1);
							if ($s->{O_RETAIL} le 0){
								$worksheet->write($a,23, "",$subt); }
							else{
								$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$subt); }
								
						$worksheet->write($a,24, "",$border1);
						$worksheet->write($a,25, "",$subt);
						$worksheet->write($a,26, "",$border1);
						$worksheet->write($a,27, "",$subt);	
						
						$worksheet->write($a,28, $s->{AMT501000},$border1);
						$worksheet->write($a,29, "",$border1);
						$worksheet->write($a,30, "",$border1);
						
						$worksheet->write($a,31, $s->{AMT503200},$border1);
						$worksheet->write($a,32, "",$border1);
						$worksheet->write($a,33, "",$border1);
						
						$worksheet->write($a,34, $s->{AMT503250},$border1);
						$worksheet->write($a,35, "",$border1);
						$worksheet->write($a,36, "",$border1);
						
						$worksheet->write($a,37, $s->{AMT503500},$border1);
						$worksheet->write($a,38, "",$border1);
						$worksheet->write($a,39, "",$border1);
						
						$worksheet->write($a,40, $s->{AMT506000},$border1);
						$worksheet->write($a,41, "",$border1);
						$worksheet->write($a,42, "",$border1);
						
						$worksheet->write($a,43, $s->{AMT503000},$border1);
						$worksheet->write($a,44, "",$border1);
						$worksheet->write($a,45, "",$border1);				
									
						$worksheet->write($a,46, $s->{AMT507000},$border1);
						$worksheet->write($a,47, "",$border1);
						$worksheet->write($a,48, "",$border1);
						
						$worksheet->write($a,49, $s->{AMT999998},$border1);
						$worksheet->write($a,50, "",$border1);
						$worksheet->write($a,51, "",$border1);
						
						$worksheet->write($a,52, $s->{AMT504000},$border1);
						$worksheet->write($a,53, "",$border1);
						$worksheet->write($a,54, "",$border1);
									
						$worksheet->write($a,55, $s->{AMT432000},$border1);
						$worksheet->write($a,56, "",$border1);
						$worksheet->write($a,57, "",$border1);
						
						$worksheet->write($a,58, $s->{AMT433000},$border1);
						$worksheet->write($a,59, "",$border1);
						$worksheet->write($a,60, "",$border1);
						
						$worksheet->write($a,61, $s->{AMT458490},$border1);
						$worksheet->write($a,62, "",$border1);
						$worksheet->write($a,63, "",$border1);
						
						$worksheet->write($a,64, $s->{AMT505000},$border1);
						$worksheet->write($a,65, "",$border1);
						$worksheet->write($a,66, "",$border1);
						
						$worksheet->write($a,67, $s->{C_RETAIL},$border1);
						$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$border1);
							if ($s->{C_RETAIL} le 0){
								$worksheet->write($a,69, "",$subt); }
							else{
								$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$subt); }
								
						$worksheet->write($a,70, "",$border1);
						$worksheet->write($a,71, "",$subt);
						$worksheet->write($a,72, "",$border1);
						$worksheet->write($a,73, "",$subt);	
						
						$worksheet->write($a,74, $s->{C_RETAIL},$border1);
						$worksheet->write($a,75, $s->{C_MARGIN},$border1);
							if ($s->{C_RETAIL} le 0){
								$worksheet->write($a,76, "",$subt); }
							else{
								$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$subt); }
							
						$worksheet->write($a,77, "",$border1);
						$worksheet->write($a,78, "",$subt);
						$worksheet->write($a,79, "",$border1);
						$worksheet->write($a,80, "",$subt);	
						
						$worksheet->write($a,81, $s->{AMT502000},$border1);
						$worksheet->write($a,82, "",$border1);
						$worksheet->write($a,83, "",$border1);
								
						$worksheet->write($a,84, $s->{AMT434000},$border1);
						$worksheet->write($a,85, "",$border1);
						$worksheet->write($a,86, "",$border1);
								
						$worksheet->write($a,87, $s->{AMT458550},$border1);
						$worksheet->write($a,88, "",$border1);
						$worksheet->write($a,89, "",$border1);
						
						$worksheet->write($a,90, $s->{AMT460100},$border1);
						$worksheet->write($a,91, "",$border1);
						$worksheet->write($a,92, "",$border1);
							
						$worksheet->write($a,93, $s->{AMT460200},$border1);
						$worksheet->write($a,94, "",$border1);
						$worksheet->write($a,95, "",$border1);
								
						$worksheet->write($a,96, $s->{AMT460300},$border1);
						$worksheet->write($a,97, "",$border1);
						$worksheet->write($a,98, "",$border1);			
						
						$a++;
						$counter++;
				
					}
					
					$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$bodyNum);
					$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
						if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
							$worksheet->write($a,9, "",$bodyPct); }
						else{
							$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$bodyPct); }
					
					$worksheet->write($a,10, "",$bodyNum);
					$worksheet->write($a,11, "",$bodyPct);
					$worksheet->write($a,12, "",$bodyNum);
					$worksheet->write($a,13, "",$bodyPct);
					
					$worksheet->write($a,14, $s->{O_RETAIL},$bodyNum);
					$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$bodyNum);
						if ($s->{O_RETAIL} le 0){
							$worksheet->write($a,16, "",$bodyPct); }
						else{
							$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$bodyPct); }
					
					$worksheet->write($a,17, "",$bodyNum);
					$worksheet->write($a,18, "",$bodyPct);
					$worksheet->write($a,19, "",$bodyNum);
					$worksheet->write($a,20, "",$bodyPct);			
					
					$worksheet->write($a,21, $s->{O_RETAIL},$bodyNum);
					$worksheet->write($a,22, $s->{O_MARGIN},$bodyNum);
						if ($s->{O_RETAIL} le 0){
							$worksheet->write($a,23, "",$bodyPct); }
						else{
							$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$bodyPct); }
							
					$worksheet->write($a,24, "",$bodyNum);
					$worksheet->write($a,25, "",$bodyPct);
					$worksheet->write($a,26, "",$bodyNum);
					$worksheet->write($a,27, "",$bodyPct);	
					
					$worksheet->write($a,28, $s->{AMT501000},$bodyNum);
					$worksheet->write($a,29, "",$bodyNum);
					$worksheet->write($a,30, "",$bodyNum);
					
					$worksheet->write($a,31, $s->{AMT503200},$bodyNum);
					$worksheet->write($a,32, "",$bodyNum);
					$worksheet->write($a,33, "",$bodyNum);
					
					$worksheet->write($a,34, $s->{AMT503250},$bodyNum);
					$worksheet->write($a,35, "",$bodyNum);
					$worksheet->write($a,36, "",$bodyNum);
					
					$worksheet->write($a,37, $s->{AMT503500},$bodyNum);
					$worksheet->write($a,38, "",$bodyNum);
					$worksheet->write($a,39, "",$bodyNum);
					
					$worksheet->write($a,40, $s->{AMT506000},$bodyNum);
					$worksheet->write($a,41, "",$bodyNum);
					$worksheet->write($a,42, "",$bodyNum);
					
					$worksheet->write($a,43, $s->{AMT503000},$bodyNum);
					$worksheet->write($a,44, "",$bodyNum);
					$worksheet->write($a,45, "",$bodyNum);				
								
					$worksheet->write($a,46, $s->{AMT507000},$bodyNum);
					$worksheet->write($a,47, "",$bodyNum);
					$worksheet->write($a,48, "",$bodyNum);
					
					$worksheet->write($a,49, $s->{AMT999998},$bodyNum);
					$worksheet->write($a,50, "",$bodyNum);
					$worksheet->write($a,51, "",$bodyNum);
					
					$worksheet->write($a,52, $s->{AMT504000},$bodyNum);
					$worksheet->write($a,53, "",$bodyNum);
					$worksheet->write($a,54, "",$bodyNum);
								
					$worksheet->write($a,55, $s->{AMT432000},$bodyNum);
					$worksheet->write($a,56, "",$bodyNum);
					$worksheet->write($a,57, "",$bodyNum);
					
					$worksheet->write($a,58, $s->{AMT433000},$bodyNum);
					$worksheet->write($a,59, "",$bodyNum);
					$worksheet->write($a,60, "",$bodyNum);
					
					$worksheet->write($a,61, $s->{AMT458490},$bodyNum);
					$worksheet->write($a,62, "",$bodyNum);
					$worksheet->write($a,63, "",$bodyNum);
					
					$worksheet->write($a,64, $s->{AMT505000},$bodyNum);
					$worksheet->write($a,65, "",$bodyNum);
					$worksheet->write($a,66, "",$bodyNum);
					
					$worksheet->write($a,67, $s->{C_RETAIL},$bodyNum);
					$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
						if ($s->{C_RETAIL} le 0){
							$worksheet->write($a,69, "",$bodyPct); }
						else{
							$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$bodyPct); }
							
					$worksheet->write($a,70, "",$bodyNum);
					$worksheet->write($a,71, "",$bodyPct);
					$worksheet->write($a,72, "",$bodyNum);
					$worksheet->write($a,73, "",$bodyPct);	
					
					$worksheet->write($a,74, $s->{C_RETAIL},$bodyNum);
					$worksheet->write($a,75, $s->{C_MARGIN},$bodyNum);
						if ($s->{C_RETAIL} le 0){
							$worksheet->write($a,76, "",$bodyPct); }
						else{
							$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$bodyPct); }
						
					$worksheet->write($a,77, "",$bodyNum);
					$worksheet->write($a,78, "",$bodyPct);
					$worksheet->write($a,79, "",$bodyNum);
					$worksheet->write($a,80, "",$bodyPct);	
					
					$worksheet->write($a,81, $s->{AMT502000},$bodyNum);
					$worksheet->write($a,82, "",$bodyNum);
					$worksheet->write($a,83, "",$bodyNum);
							
					$worksheet->write($a,84, $s->{AMT434000},$bodyNum);
					$worksheet->write($a,85, "",$bodyNum);
					$worksheet->write($a,86, "",$bodyNum);
							
					$worksheet->write($a,87, $s->{AMT458550},$bodyNum);
					$worksheet->write($a,88, "",$bodyNum);
					$worksheet->write($a,89, "",$bodyNum);
					
					$worksheet->write($a,90, $s->{AMT460100},$bodyNum);
					$worksheet->write($a,91, "",$bodyNum);
					$worksheet->write($a,92, "",$bodyNum);
						
					$worksheet->write($a,93, $s->{AMT460200},$bodyNum);
					$worksheet->write($a,94, "",$bodyNum);
					$worksheet->write($a,95, "",$bodyNum);
							
					$worksheet->write($a,96, $s->{AMT460300},$bodyNum);
					$worksheet->write($a,97, "",$bodyNum);
					$worksheet->write($a,98, "",$bodyNum);

					$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
					$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
					
					$counter = 0;
					$a++;
				}
				
				$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$bodyNum);
				$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
					if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
						$worksheet->write($a,9, "",$bodyPct); }
					else{
						$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$bodyPct); }
				
				$worksheet->write($a,10, "",$bodyNum);
				$worksheet->write($a,11, "",$bodyPct);
				$worksheet->write($a,12, "",$bodyNum);
				$worksheet->write($a,13, "",$bodyPct);
				
				$worksheet->write($a,14, $s->{O_RETAIL},$bodyNum);
				$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$bodyNum);
					if ($s->{O_RETAIL} le 0){
						$worksheet->write($a,16, "",$bodyPct); }
					else{
						$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$bodyPct); }
				
				$worksheet->write($a,17, "",$bodyNum);
				$worksheet->write($a,18, "",$bodyPct);
				$worksheet->write($a,19, "",$bodyNum);
				$worksheet->write($a,20, "",$bodyPct);			
				
				$worksheet->write($a,21, $s->{O_RETAIL},$bodyNum);
				$worksheet->write($a,22, $s->{O_MARGIN},$bodyNum);
					if ($s->{O_RETAIL} le 0){
						$worksheet->write($a,23, "",$bodyPct); }
					else{
						$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$bodyPct); }
						
				$worksheet->write($a,24, "",$bodyNum);
				$worksheet->write($a,25, "",$bodyPct);
				$worksheet->write($a,26, "",$bodyNum);
				$worksheet->write($a,27, "",$bodyPct);	
				
				$worksheet->write($a,28, $s->{AMT501000},$bodyNum);
				$worksheet->write($a,29, "",$bodyNum);
				$worksheet->write($a,30, "",$bodyNum);
				
				$worksheet->write($a,31, $s->{AMT503200},$bodyNum);
				$worksheet->write($a,32, "",$bodyNum);
				$worksheet->write($a,33, "",$bodyNum);
				
				$worksheet->write($a,34, $s->{AMT503250},$bodyNum);
				$worksheet->write($a,35, "",$bodyNum);
				$worksheet->write($a,36, "",$bodyNum);
				
				$worksheet->write($a,37, $s->{AMT503500},$bodyNum);
				$worksheet->write($a,38, "",$bodyNum);
				$worksheet->write($a,39, "",$bodyNum);
				
				$worksheet->write($a,40, $s->{AMT506000},$bodyNum);
				$worksheet->write($a,41, "",$bodyNum);
				$worksheet->write($a,42, "",$bodyNum);
				
				$worksheet->write($a,43, $s->{AMT503000},$bodyNum);
				$worksheet->write($a,44, "",$bodyNum);
				$worksheet->write($a,45, "",$bodyNum);				
							
				$worksheet->write($a,46, $s->{AMT507000},$bodyNum);
				$worksheet->write($a,47, "",$bodyNum);
				$worksheet->write($a,48, "",$bodyNum);
				
				$worksheet->write($a,49, $s->{AMT999998},$bodyNum);
				$worksheet->write($a,50, "",$bodyNum);
				$worksheet->write($a,51, "",$bodyNum);
				
				$worksheet->write($a,52, $s->{AMT504000},$bodyNum);
				$worksheet->write($a,53, "",$bodyNum);
				$worksheet->write($a,54, "",$bodyNum);
							
				$worksheet->write($a,55, $s->{AMT432000},$bodyNum);
				$worksheet->write($a,56, "",$bodyNum);
				$worksheet->write($a,57, "",$bodyNum);
				
				$worksheet->write($a,58, $s->{AMT433000},$bodyNum);
				$worksheet->write($a,59, "",$bodyNum);
				$worksheet->write($a,60, "",$bodyNum);
				
				$worksheet->write($a,61, $s->{AMT458490},$bodyNum);
				$worksheet->write($a,62, "",$bodyNum);
				$worksheet->write($a,63, "",$bodyNum);
				
				$worksheet->write($a,64, $s->{AMT505000},$bodyNum);
				$worksheet->write($a,65, "",$bodyNum);
				$worksheet->write($a,66, "",$bodyNum);
				
				$worksheet->write($a,67, $s->{C_RETAIL},$bodyNum);
				$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
					if ($s->{C_RETAIL} le 0){
						$worksheet->write($a,69, "",$bodyPct); }
					else{
						$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$bodyPct); }
						
				$worksheet->write($a,70, "",$bodyNum);
				$worksheet->write($a,71, "",$bodyPct);
				$worksheet->write($a,72, "",$bodyNum);
				$worksheet->write($a,73, "",$bodyPct);	
				
				$worksheet->write($a,74, $s->{C_RETAIL},$bodyNum);
				$worksheet->write($a,75, $s->{C_MARGIN},$bodyNum);
					if ($s->{C_RETAIL} le 0){
						$worksheet->write($a,76, "",$bodyPct); }
					else{
						$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$bodyPct); }
					
				$worksheet->write($a,77, "",$bodyNum);
				$worksheet->write($a,78, "",$bodyPct);
				$worksheet->write($a,79, "",$bodyNum);
				$worksheet->write($a,80, "",$bodyPct);	
				
				$worksheet->write($a,81, $s->{AMT502000},$bodyNum);
				$worksheet->write($a,82, "",$bodyNum);
				$worksheet->write($a,83, "",$bodyNum);
						
				$worksheet->write($a,84, $s->{AMT434000},$bodyNum);
				$worksheet->write($a,85, "",$bodyNum);
				$worksheet->write($a,86, "",$bodyNum);
						
				$worksheet->write($a,87, $s->{AMT458550},$bodyNum);
				$worksheet->write($a,88, "",$bodyNum);
				$worksheet->write($a,89, "",$bodyNum);
				
				$worksheet->write($a,90, $s->{AMT460100},$bodyNum);
				$worksheet->write($a,91, "",$bodyNum);
				$worksheet->write($a,92, "",$bodyNum);
					
				$worksheet->write($a,93, $s->{AMT460200},$bodyNum);
				$worksheet->write($a,94, "",$bodyNum);
				$worksheet->write($a,95, "",$bodyNum);
						
				$worksheet->write($a,96, $s->{AMT460300},$bodyNum);
				$worksheet->write($a,97, "",$bodyNum);
				$worksheet->write($a,98, "",$bodyNum);

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
			
			$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$headNumber);
			$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$headNumber);
				if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
					$worksheet->write($a,9, "",$headPct); }
				else{
					$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$headPct); }
				
			$worksheet->write($a,10, "",$headNumber);
			$worksheet->write($a,11, "",$headPct);
			$worksheet->write($a,12, "",$headNumber);
			$worksheet->write($a,13, "",$headPct);
			
			$worksheet->write($a,14, $s->{O_RETAIL},$headNumber);
			$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$headNumber);
				if ($s->{O_RETAIL} le 0){
					$worksheet->write($a,16, "",$headPct); }
				else{
					$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$headPct); }
				
			$worksheet->write($a,17, "",$headNumber);
			$worksheet->write($a,18, "",$headPct);
			$worksheet->write($a,19, "",$headNumber);
			$worksheet->write($a,20, "",$headPct);			
				
			$worksheet->write($a,21, $s->{O_RETAIL},$headNumber);
			$worksheet->write($a,22, $s->{O_MARGIN},$headNumber);
				if ($s->{O_RETAIL} le 0){
					$worksheet->write($a,23, "",$headPct); }
				else{
					$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$headPct); }
						
			$worksheet->write($a,24, "",$headNumber);
			$worksheet->write($a,25, "",$headPct);
			$worksheet->write($a,26, "",$headNumber);
			$worksheet->write($a,27, "",$headPct);	
				
			$worksheet->write($a,28, $s->{AMT501000},$headNumber);
			$worksheet->write($a,29, "",$headNumber);
			$worksheet->write($a,30, "",$headNumber);
				
			$worksheet->write($a,31, $s->{AMT503200},$headNumber);
			$worksheet->write($a,32, "",$headNumber);
			$worksheet->write($a,33, "",$headNumber);
				
			$worksheet->write($a,34, $s->{AMT503250},$headNumber);
			$worksheet->write($a,35, "",$headNumber);
			$worksheet->write($a,36, "",$headNumber);
				
			$worksheet->write($a,37, $s->{AMT503500},$headNumber);
			$worksheet->write($a,38, "",$headNumber);
			$worksheet->write($a,39, "",$headNumber);
				
			$worksheet->write($a,40, $s->{AMT506000},$headNumber);
			$worksheet->write($a,41, "",$headNumber);
			$worksheet->write($a,42, "",$headNumber);
			
			$worksheet->write($a,43, $s->{AMT503000},$headNumber);
			$worksheet->write($a,44, "",$headNumber);
			$worksheet->write($a,45, "",$headNumber);				
							
			$worksheet->write($a,46, $s->{AMT507000},$headNumber);
			$worksheet->write($a,47, "",$headNumber);
			$worksheet->write($a,48, "",$headNumber);
				
			$worksheet->write($a,49, $s->{AMT999998},$headNumber);
			$worksheet->write($a,50, "",$headNumber);
			$worksheet->write($a,51, "",$headNumber);
				
			$worksheet->write($a,52, $s->{AMT504000},$headNumber);
			$worksheet->write($a,53, "",$headNumber);
			$worksheet->write($a,54, "",$headNumber);
							
			$worksheet->write($a,55, $s->{AMT432000},$headNumber);
			$worksheet->write($a,56, "",$headNumber);
			$worksheet->write($a,57, "",$headNumber);
				
			$worksheet->write($a,58, $s->{AMT433000},$headNumber);
			$worksheet->write($a,59, "",$headNumber);
			$worksheet->write($a,60, "",$headNumber);
				
			$worksheet->write($a,61, $s->{AMT458490},$headNumber);
			$worksheet->write($a,62, "",$headNumber);
			$worksheet->write($a,63, "",$headNumber);
			
			$worksheet->write($a,64, $s->{AMT505000},$headNumber);
			$worksheet->write($a,65, "",$headNumber);
			$worksheet->write($a,66, "",$headNumber);
			
			$worksheet->write($a,67, $s->{C_RETAIL},$headNumber);
			$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$headNumber);
				if ($s->{C_RETAIL} le 0){
					$worksheet->write($a,69, "",$headPct); }
				else{
					$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$headPct); }						
			
			$worksheet->write($a,70, "",$headNumber);
			$worksheet->write($a,71, "",$headPct);
			$worksheet->write($a,72, "",$headNumber);
			$worksheet->write($a,73, "",$headPct);	
				
			$worksheet->write($a,74, $s->{C_RETAIL},$headNumber);
			$worksheet->write($a,75, $s->{C_MARGIN},$headNumber);
				if ($s->{C_RETAIL} le 0){
					$worksheet->write($a,76, "",$headPct); }
				else{
					$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$headPct); }
				
			$worksheet->write($a,77, "",$headNumber);
			$worksheet->write($a,78, "",$headPct);
			$worksheet->write($a,79, "",$headNumber);
			$worksheet->write($a,80, "",$headPct);	
				
			$worksheet->write($a,81, $s->{AMT502000},$headNumber);
			$worksheet->write($a,82, "",$headNumber);
			$worksheet->write($a,83, "",$headNumber);
						
			$worksheet->write($a,84, $s->{AMT434000},$headNumber);
			$worksheet->write($a,85, "",$headNumber);
			$worksheet->write($a,86, "",$headNumber);
					
			$worksheet->write($a,87, $s->{AMT458550},$headNumber);
			$worksheet->write($a,88, "",$headNumber);
			$worksheet->write($a,89, "",$headNumber);
				
			$worksheet->write($a,90, $s->{AMT460100},$headNumber);
			$worksheet->write($a,91, "",$headNumber);
			$worksheet->write($a,92, "",$headNumber);
					
			$worksheet->write($a,93, $s->{AMT460200},$headNumber);
			$worksheet->write($a,94, "",$headNumber);
			$worksheet->write($a,95, "",$headNumber);
						
			$worksheet->write($a,96, $s->{AMT460300},$headNumber);
			$worksheet->write($a,97, "",$headNumber);
			$worksheet->write($a,98, "",$headNumber);
			
			$a++;
		}

	$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$headNumber);
	$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$headNumber);
		if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
			$worksheet->write($a,9, "",$headPct); }
		else{
			$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$headPct); }
				
	$worksheet->write($a,10, "",$headNumber);
	$worksheet->write($a,11, "",$headPct);
	$worksheet->write($a,12, "",$headNumber);
	$worksheet->write($a,13, "",$headPct);
		
	$worksheet->write($a,14, $s->{O_RETAIL},$headNumber);
	$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$headNumber);
		if ($s->{O_RETAIL} le 0){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$headPct); }
				
	$worksheet->write($a,17, "",$headNumber);
	$worksheet->write($a,18, "",$headPct);
	$worksheet->write($a,19, "",$headNumber);
	$worksheet->write($a,20, "",$headPct);			
	
	$worksheet->write($a,21, $s->{O_RETAIL},$headNumber);
	$worksheet->write($a,22, $s->{O_MARGIN},$headNumber);
		if ($s->{O_RETAIL} le 0){
			$worksheet->write($a,23, "",$headPct); }
		else{
			$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$headPct); }
						
	$worksheet->write($a,24, "",$headNumber);
	$worksheet->write($a,25, "",$headPct);
	$worksheet->write($a,26, "",$headNumber);
	$worksheet->write($a,27, "",$headPct);	
				
	$worksheet->write($a,28, $s->{AMT501000},$headNumber);
	$worksheet->write($a,29, "",$headNumber);
	$worksheet->write($a,30, "",$headNumber);
				
	$worksheet->write($a,31, $s->{AMT503200},$headNumber);
	$worksheet->write($a,32, "",$headNumber);
	$worksheet->write($a,33, "",$headNumber);
				
	$worksheet->write($a,34, $s->{AMT503250},$headNumber);
	$worksheet->write($a,35, "",$headNumber);
	$worksheet->write($a,36, "",$headNumber);
				
	$worksheet->write($a,37, $s->{AMT503500},$headNumber);
	$worksheet->write($a,38, "",$headNumber);
	$worksheet->write($a,39, "",$headNumber);
			
	$worksheet->write($a,40, $s->{AMT506000},$headNumber);
	$worksheet->write($a,41, "",$headNumber);
	$worksheet->write($a,42, "",$headNumber);
				
	$worksheet->write($a,43, $s->{AMT503000},$headNumber);
	$worksheet->write($a,44, "",$headNumber);
	$worksheet->write($a,45, "",$headNumber);				
							
	$worksheet->write($a,46, $s->{AMT507000},$headNumber);
	$worksheet->write($a,47, "",$headNumber);
	$worksheet->write($a,48, "",$headNumber);
				
	$worksheet->write($a,49, $s->{AMT999998},$headNumber);
	$worksheet->write($a,50, "",$headNumber);
	$worksheet->write($a,51, "",$headNumber);
				
	$worksheet->write($a,52, $s->{AMT504000},$headNumber);
	$worksheet->write($a,53, "",$headNumber);
	$worksheet->write($a,54, "",$headNumber);
							
	$worksheet->write($a,55, $s->{AMT432000},$headNumber);
	$worksheet->write($a,56, "",$headNumber);
	$worksheet->write($a,57, "",$headNumber);
				
	$worksheet->write($a,58, $s->{AMT433000},$headNumber);
	$worksheet->write($a,59, "",$headNumber);
	$worksheet->write($a,60, "",$headNumber);
	
	$worksheet->write($a,61, $s->{AMT458490},$headNumber);
	$worksheet->write($a,62, "",$headNumber);
	$worksheet->write($a,63, "",$headNumber);
				
	$worksheet->write($a,64, $s->{AMT505000},$headNumber);
	$worksheet->write($a,65, "",$headNumber);
	$worksheet->write($a,66, "",$headNumber);
				
	$worksheet->write($a,67, $s->{C_RETAIL},$headNumber);
	$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$headNumber);
		if ($s->{C_RETAIL} le 0){
			$worksheet->write($a,69, "",$headPct); }
		else{
			$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$headPct); }
						
	$worksheet->write($a,70, "",$headNumber);
	$worksheet->write($a,71, "",$headPct);
	$worksheet->write($a,72, "",$headNumber);
	$worksheet->write($a,73, "",$headPct);	
				
	$worksheet->write($a,74, $s->{C_RETAIL},$headNumber);
	$worksheet->write($a,75, $s->{C_MARGIN},$headNumber);
		if ($s->{C_RETAIL} le 0){
			$worksheet->write($a,76, "",$headPct); }
		else{
			$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$headPct); }
		
	$worksheet->write($a,77, "",$headNumber);
	$worksheet->write($a,78, "",$headPct);
	$worksheet->write($a,79, "",$headNumber);
	$worksheet->write($a,80, "",$headPct);	
				
	$worksheet->write($a,81, $s->{AMT502000},$headNumber);
	$worksheet->write($a,82, "",$headNumber);
	$worksheet->write($a,83, "",$headNumber);
						
	$worksheet->write($a,84, $s->{AMT434000},$headNumber);
	$worksheet->write($a,85, "",$headNumber);
	$worksheet->write($a,86, "",$headNumber);
					
	$worksheet->write($a,87, $s->{AMT458550},$headNumber);
	$worksheet->write($a,88, "",$headNumber);
	$worksheet->write($a,89, "",$headNumber);
	
	$worksheet->write($a,90, $s->{AMT460100},$headNumber);
	$worksheet->write($a,91, "",$headNumber);
	$worksheet->write($a,92, "",$headNumber);
			
	$worksheet->write($a,93, $s->{AMT460200},$headNumber);
	$worksheet->write($a,94, "",$headNumber);
	$worksheet->write($a,95, "",$headNumber);
						
	$worksheet->write($a,96, $s->{AMT460300},$headNumber);
	$worksheet->write($a,97, "",$headNumber);
	$worksheet->write($a,98, "",$headNumber);

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
			SELECT SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
			FROM METRO_IT_MARGIN_DEPT
			WHERE STORE_CODE = '$store'
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
			});								 
$sls->execute();

	while(my $s = $sls->fetchrow_hashref()){
				
	$sls1 = $dbh->prepare (qq{
				SELECT STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
				FROM METRO_IT_MARGIN_DEPT
				WHERE STORE_CODE = '$store'
					AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
				GROUP BY STORE_CODE, STORE_DESCRIPTION, MERCH_GROUP_CODE, MERCH_GROUP_DESC
				ORDER BY 1, 3
				});								 
	$sls1->execute();

		while(my $s = $sls1->fetchrow_hashref()){
			$merch_group_code = $s->{MERCH_GROUP_CODE};
			$merch_group_desc = $s->{MERCH_GROUP_DESC}; 
			$loc_code = $s->{STORE_CODE};
			$loc_desc = $s->{STORE_DESCRIPTION};
			
			$sls2 = $dbh->prepare (qq{
					SELECT GROUP_CODE, GROUP_DESC, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
					FROM METRO_IT_MARGIN_DEPT
					WHERE MERCH_GROUP_CODE = '$merch_group_code' AND STORE_CODE = '$store'
						AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
					GROUP BY GROUP_CODE, GROUP_DESC
					ORDER BY 1
					});	
			$sls2->execute();
			
			$mgc_counter = $a;
			while(my $s = $sls2->fetchrow_hashref()){
				$group_code = $s->{GROUP_CODE};
				$group_desc = $s->{GROUP_DESC};
						
				$sls3 = $dbh->prepare (qq{
						SELECT DIVISION, DIVISION_DESC, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
						FROM METRO_IT_MARGIN_DEPT
						WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' AND STORE_CODE = '$store'
							AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
						GROUP BY DIVISION, DIVISION_DESC
						ORDER BY 1
					});
				$sls3->execute();
				
				$grp_counter = $a;
				while(my $s = $sls3->fetchrow_hashref()){
					$division = $s->{DIVISION};
					$division_desc = $s->{DIVISION_DESC};
					
					$sls4 = $dbh->prepare (qq{	 
							SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, SUM(OTR_COST) AS O_COST, SUM(OTR_RETAIL) AS O_RETAIL, SUM(OTR_MARGIN) AS O_MARGIN, SUM(CON_COST) AS C_COST, SUM(CON_RETAIL) AS C_RETAIL, SUM(CON_MARGIN) AS C_MARGIN, SUM(AMOUNT432000) AS AMT432000, SUM(AMOUNT433000) AS AMT433000, SUM(AMOUNT458490) AS AMT458490, SUM(AMOUNT434000) AS AMT434000, SUM(AMOUNT458550) AS AMT458550, SUM(AMOUNT460100) AS AMT460100, SUM(AMOUNT460200) AS AMT460200, SUM(AMOUNT460300) AS AMT460300 , SUM(AMOUNT503200) AS AMT503200 , SUM(AMOUNT503250) AS AMT503250, SUM(AMOUNT503500) AS AMT503500, SUM(AMOUNT506000) AS AMT506000, SUM(AMOUNT501000) AS AMT501000, SUM(AMOUNT503000) AS AMT503000, SUM(AMOUNT507000) AS AMT507000, SUM(AMOUNT999998) AS AMT999998, SUM(AMOUNT505000) AS AMT505000, SUM(AMOUNT504000) AS AMT504000, SUM(AMOUNT502000) AS AMT502000
							 FROM METRO_IT_MARGIN_DEPT
							 WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
								AND STORE_CODE = '$store'
								AND YEAR = (SELECT TO_CHAR(SYSDATE, 'YYYY') FROM DUAL) AND (PER = UPPER(TO_CHAR(TO_DATE('$mo_st_date_fld'), 'MON')))
							 GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC
							 ORDER BY 1
						});
					$sls4->execute();
					
					while(my $s = $sls4->fetchrow_hashref()){
						
						$worksheet->write($a,5, $s->{DEPARTMENT_CODE},$desc);
						$worksheet->write($a,6, $s->{DEPARTMENT_DESC},$desc);
						
						$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$border1);
						$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$border1);
							if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
								$worksheet->write($a,9, "",$subt); }
							else{
								$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$subt); }
						
						$worksheet->write($a,10, "",$border1);
						$worksheet->write($a,11, "",$subt);
						$worksheet->write($a,12, "",$border1);
						$worksheet->write($a,13, "",$subt);
						
						$worksheet->write($a,14, $s->{O_RETAIL},$border1);
						$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$border1);
							if ($s->{O_RETAIL} le 0){
								$worksheet->write($a,16, "",$subt); }
							else{
								$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$subt); }
						
						$worksheet->write($a,17, "",$border1);
						$worksheet->write($a,18, "",$subt);
						$worksheet->write($a,19, "",$border1);
						$worksheet->write($a,20, "",$subt);			
						
						$worksheet->write($a,21, $s->{O_RETAIL},$border1);
						$worksheet->write($a,22, $s->{O_MARGIN},$border1);
							if ($s->{O_RETAIL} le 0){
								$worksheet->write($a,23, "",$subt); }
							else{
								$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$subt); }
								
						$worksheet->write($a,24, "",$border1);
						$worksheet->write($a,25, "",$subt);
						$worksheet->write($a,26, "",$border1);
						$worksheet->write($a,27, "",$subt);	
						
						$worksheet->write($a,28, $s->{AMT501000},$border1);
						$worksheet->write($a,29, "",$border1);
						$worksheet->write($a,30, "",$border1);
						
						$worksheet->write($a,31, $s->{AMT503200},$border1);
						$worksheet->write($a,32, "",$border1);
						$worksheet->write($a,33, "",$border1);
						
						$worksheet->write($a,34, $s->{AMT503250},$border1);
						$worksheet->write($a,35, "",$border1);
						$worksheet->write($a,36, "",$border1);
						
						$worksheet->write($a,37, $s->{AMT503500},$border1);
						$worksheet->write($a,38, "",$border1);
						$worksheet->write($a,39, "",$border1);
						
						$worksheet->write($a,40, $s->{AMT506000},$border1);
						$worksheet->write($a,41, "",$border1);
						$worksheet->write($a,42, "",$border1);
						
						$worksheet->write($a,43, $s->{AMT503000},$border1);
						$worksheet->write($a,44, "",$border1);
						$worksheet->write($a,45, "",$border1);				
									
						$worksheet->write($a,46, $s->{AMT507000},$border1);
						$worksheet->write($a,47, "",$border1);
						$worksheet->write($a,48, "",$border1);
						
						$worksheet->write($a,49, $s->{AMT999998},$border1);
						$worksheet->write($a,50, "",$border1);
						$worksheet->write($a,51, "",$border1);
						
						$worksheet->write($a,52, $s->{AMT504000},$border1);
						$worksheet->write($a,53, "",$border1);
						$worksheet->write($a,54, "",$border1);
									
						$worksheet->write($a,55, $s->{AMT432000},$border1);
						$worksheet->write($a,56, "",$border1);
						$worksheet->write($a,57, "",$border1);
						
						$worksheet->write($a,58, $s->{AMT433000},$border1);
						$worksheet->write($a,59, "",$border1);
						$worksheet->write($a,60, "",$border1);
						
						$worksheet->write($a,61, $s->{AMT458490},$border1);
						$worksheet->write($a,62, "",$border1);
						$worksheet->write($a,63, "",$border1);
						
						$worksheet->write($a,64, $s->{AMT505000},$border1);
						$worksheet->write($a,65, "",$border1);
						$worksheet->write($a,66, "",$border1);
						
						$worksheet->write($a,67, $s->{C_RETAIL},$border1);
						$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$border1);
							if ($s->{C_RETAIL} le 0){
								$worksheet->write($a,69, "",$subt); }
							else{
								$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$subt); }
								
						$worksheet->write($a,70, "",$border1);
						$worksheet->write($a,71, "",$subt);
						$worksheet->write($a,72, "",$border1);
						$worksheet->write($a,73, "",$subt);	
						
						$worksheet->write($a,74, $s->{C_RETAIL},$border1);
						$worksheet->write($a,75, $s->{C_MARGIN},$border1);
							if ($s->{C_RETAIL} le 0){
								$worksheet->write($a,76, "",$subt); }
							else{
								$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$subt); }
							
						$worksheet->write($a,77, "",$border1);
						$worksheet->write($a,78, "",$subt);
						$worksheet->write($a,79, "",$border1);
						$worksheet->write($a,80, "",$subt);	
						
						$worksheet->write($a,81, $s->{AMT502000},$border1);
						$worksheet->write($a,82, "",$border1);
						$worksheet->write($a,83, "",$border1);
								
						$worksheet->write($a,84, $s->{AMT434000},$border1);
						$worksheet->write($a,85, "",$border1);
						$worksheet->write($a,86, "",$border1);
								
						$worksheet->write($a,87, $s->{AMT458550},$border1);
						$worksheet->write($a,88, "",$border1);
						$worksheet->write($a,89, "",$border1);
						
						$worksheet->write($a,90, $s->{AMT460100},$border1);
						$worksheet->write($a,91, "",$border1);
						$worksheet->write($a,92, "",$border1);
							
						$worksheet->write($a,93, $s->{AMT460200},$border1);
						$worksheet->write($a,94, "",$border1);
						$worksheet->write($a,95, "",$border1);
								
						$worksheet->write($a,96, $s->{AMT460300},$border1);
						$worksheet->write($a,97, "",$border1);
						$worksheet->write($a,98, "",$border1);			
						
						$a++;
						$counter++;
				
					}
					
					$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$bodyNum);
					$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
						if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
							$worksheet->write($a,9, "",$bodyPct); }
						else{
							$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$bodyPct); }
					
					$worksheet->write($a,10, "",$bodyNum);
					$worksheet->write($a,11, "",$bodyPct);
					$worksheet->write($a,12, "",$bodyNum);
					$worksheet->write($a,13, "",$bodyPct);
					
					$worksheet->write($a,14, $s->{O_RETAIL},$bodyNum);
					$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$bodyNum);
						if ($s->{O_RETAIL} le 0){
							$worksheet->write($a,16, "",$bodyPct); }
						else{
							$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$bodyPct); }
					
					$worksheet->write($a,17, "",$bodyNum);
					$worksheet->write($a,18, "",$bodyPct);
					$worksheet->write($a,19, "",$bodyNum);
					$worksheet->write($a,20, "",$bodyPct);			
					
					$worksheet->write($a,21, $s->{O_RETAIL},$bodyNum);
					$worksheet->write($a,22, $s->{O_MARGIN},$bodyNum);
						if ($s->{O_RETAIL} le 0){
							$worksheet->write($a,23, "",$bodyPct); }
						else{
							$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$bodyPct); }
							
					$worksheet->write($a,24, "",$bodyNum);
					$worksheet->write($a,25, "",$bodyPct);
					$worksheet->write($a,26, "",$bodyNum);
					$worksheet->write($a,27, "",$bodyPct);	
					
					$worksheet->write($a,28, $s->{AMT501000},$bodyNum);
					$worksheet->write($a,29, "",$bodyNum);
					$worksheet->write($a,30, "",$bodyNum);
					
					$worksheet->write($a,31, $s->{AMT503200},$bodyNum);
					$worksheet->write($a,32, "",$bodyNum);
					$worksheet->write($a,33, "",$bodyNum);
					
					$worksheet->write($a,34, $s->{AMT503250},$bodyNum);
					$worksheet->write($a,35, "",$bodyNum);
					$worksheet->write($a,36, "",$bodyNum);
					
					$worksheet->write($a,37, $s->{AMT503500},$bodyNum);
					$worksheet->write($a,38, "",$bodyNum);
					$worksheet->write($a,39, "",$bodyNum);
					
					$worksheet->write($a,40, $s->{AMT506000},$bodyNum);
					$worksheet->write($a,41, "",$bodyNum);
					$worksheet->write($a,42, "",$bodyNum);
					
					$worksheet->write($a,43, $s->{AMT503000},$bodyNum);
					$worksheet->write($a,44, "",$bodyNum);
					$worksheet->write($a,45, "",$bodyNum);				
								
					$worksheet->write($a,46, $s->{AMT507000},$bodyNum);
					$worksheet->write($a,47, "",$bodyNum);
					$worksheet->write($a,48, "",$bodyNum);
					
					$worksheet->write($a,49, $s->{AMT999998},$bodyNum);
					$worksheet->write($a,50, "",$bodyNum);
					$worksheet->write($a,51, "",$bodyNum);
					
					$worksheet->write($a,52, $s->{AMT504000},$bodyNum);
					$worksheet->write($a,53, "",$bodyNum);
					$worksheet->write($a,54, "",$bodyNum);
								
					$worksheet->write($a,55, $s->{AMT432000},$bodyNum);
					$worksheet->write($a,56, "",$bodyNum);
					$worksheet->write($a,57, "",$bodyNum);
					
					$worksheet->write($a,58, $s->{AMT433000},$bodyNum);
					$worksheet->write($a,59, "",$bodyNum);
					$worksheet->write($a,60, "",$bodyNum);
					
					$worksheet->write($a,61, $s->{AMT458490},$bodyNum);
					$worksheet->write($a,62, "",$bodyNum);
					$worksheet->write($a,63, "",$bodyNum);
					
					$worksheet->write($a,64, $s->{AMT505000},$bodyNum);
					$worksheet->write($a,65, "",$bodyNum);
					$worksheet->write($a,66, "",$bodyNum);
					
					$worksheet->write($a,67, $s->{C_RETAIL},$bodyNum);
					$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
						if ($s->{C_RETAIL} le 0){
							$worksheet->write($a,69, "",$bodyPct); }
						else{
							$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$bodyPct); }
							
					$worksheet->write($a,70, "",$bodyNum);
					$worksheet->write($a,71, "",$bodyPct);
					$worksheet->write($a,72, "",$bodyNum);
					$worksheet->write($a,73, "",$bodyPct);	
					
					$worksheet->write($a,74, $s->{C_RETAIL},$bodyNum);
					$worksheet->write($a,75, $s->{C_MARGIN},$bodyNum);
						if ($s->{C_RETAIL} le 0){
							$worksheet->write($a,76, "",$bodyPct); }
						else{
							$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$bodyPct); }
						
					$worksheet->write($a,77, "",$bodyNum);
					$worksheet->write($a,78, "",$bodyPct);
					$worksheet->write($a,79, "",$bodyNum);
					$worksheet->write($a,80, "",$bodyPct);	
					
					$worksheet->write($a,81, $s->{AMT502000},$bodyNum);
					$worksheet->write($a,82, "",$bodyNum);
					$worksheet->write($a,83, "",$bodyNum);
							
					$worksheet->write($a,84, $s->{AMT434000},$bodyNum);
					$worksheet->write($a,85, "",$bodyNum);
					$worksheet->write($a,86, "",$bodyNum);
							
					$worksheet->write($a,87, $s->{AMT458550},$bodyNum);
					$worksheet->write($a,88, "",$bodyNum);
					$worksheet->write($a,89, "",$bodyNum);
					
					$worksheet->write($a,90, $s->{AMT460100},$bodyNum);
					$worksheet->write($a,91, "",$bodyNum);
					$worksheet->write($a,92, "",$bodyNum);
						
					$worksheet->write($a,93, $s->{AMT460200},$bodyNum);
					$worksheet->write($a,94, "",$bodyNum);
					$worksheet->write($a,95, "",$bodyNum);
							
					$worksheet->write($a,96, $s->{AMT460300},$bodyNum);
					$worksheet->write($a,97, "",$bodyNum);
					$worksheet->write($a,98, "",$bodyNum);

					$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
					$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
					
					$counter = 0;
					$a++;
				}
				
				$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$bodyNum);
				$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
					if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
						$worksheet->write($a,9, "",$bodyPct); }
					else{
						$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$bodyPct); }
				
				$worksheet->write($a,10, "",$bodyNum);
				$worksheet->write($a,11, "",$bodyPct);
				$worksheet->write($a,12, "",$bodyNum);
				$worksheet->write($a,13, "",$bodyPct);
				
				$worksheet->write($a,14, $s->{O_RETAIL},$bodyNum);
				$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$bodyNum);
					if ($s->{O_RETAIL} le 0){
						$worksheet->write($a,16, "",$bodyPct); }
					else{
						$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$bodyPct); }
				
				$worksheet->write($a,17, "",$bodyNum);
				$worksheet->write($a,18, "",$bodyPct);
				$worksheet->write($a,19, "",$bodyNum);
				$worksheet->write($a,20, "",$bodyPct);			
				
				$worksheet->write($a,21, $s->{O_RETAIL},$bodyNum);
				$worksheet->write($a,22, $s->{O_MARGIN},$bodyNum);
					if ($s->{O_RETAIL} le 0){
						$worksheet->write($a,23, "",$bodyPct); }
					else{
						$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$bodyPct); }
						
				$worksheet->write($a,24, "",$bodyNum);
				$worksheet->write($a,25, "",$bodyPct);
				$worksheet->write($a,26, "",$bodyNum);
				$worksheet->write($a,27, "",$bodyPct);	
				
				$worksheet->write($a,28, $s->{AMT501000},$bodyNum);
				$worksheet->write($a,29, "",$bodyNum);
				$worksheet->write($a,30, "",$bodyNum);
				
				$worksheet->write($a,31, $s->{AMT503200},$bodyNum);
				$worksheet->write($a,32, "",$bodyNum);
				$worksheet->write($a,33, "",$bodyNum);
				
				$worksheet->write($a,34, $s->{AMT503250},$bodyNum);
				$worksheet->write($a,35, "",$bodyNum);
				$worksheet->write($a,36, "",$bodyNum);
				
				$worksheet->write($a,37, $s->{AMT503500},$bodyNum);
				$worksheet->write($a,38, "",$bodyNum);
				$worksheet->write($a,39, "",$bodyNum);
				
				$worksheet->write($a,40, $s->{AMT506000},$bodyNum);
				$worksheet->write($a,41, "",$bodyNum);
				$worksheet->write($a,42, "",$bodyNum);
				
				$worksheet->write($a,43, $s->{AMT503000},$bodyNum);
				$worksheet->write($a,44, "",$bodyNum);
				$worksheet->write($a,45, "",$bodyNum);				
							
				$worksheet->write($a,46, $s->{AMT507000},$bodyNum);
				$worksheet->write($a,47, "",$bodyNum);
				$worksheet->write($a,48, "",$bodyNum);
				
				$worksheet->write($a,49, $s->{AMT999998},$bodyNum);
				$worksheet->write($a,50, "",$bodyNum);
				$worksheet->write($a,51, "",$bodyNum);
				
				$worksheet->write($a,52, $s->{AMT504000},$bodyNum);
				$worksheet->write($a,53, "",$bodyNum);
				$worksheet->write($a,54, "",$bodyNum);
							
				$worksheet->write($a,55, $s->{AMT432000},$bodyNum);
				$worksheet->write($a,56, "",$bodyNum);
				$worksheet->write($a,57, "",$bodyNum);
				
				$worksheet->write($a,58, $s->{AMT433000},$bodyNum);
				$worksheet->write($a,59, "",$bodyNum);
				$worksheet->write($a,60, "",$bodyNum);
				
				$worksheet->write($a,61, $s->{AMT458490},$bodyNum);
				$worksheet->write($a,62, "",$bodyNum);
				$worksheet->write($a,63, "",$bodyNum);
				
				$worksheet->write($a,64, $s->{AMT505000},$bodyNum);
				$worksheet->write($a,65, "",$bodyNum);
				$worksheet->write($a,66, "",$bodyNum);
				
				$worksheet->write($a,67, $s->{C_RETAIL},$bodyNum);
				$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$bodyNum);
					if ($s->{C_RETAIL} le 0){
						$worksheet->write($a,69, "",$bodyPct); }
					else{
						$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$bodyPct); }
						
				$worksheet->write($a,70, "",$bodyNum);
				$worksheet->write($a,71, "",$bodyPct);
				$worksheet->write($a,72, "",$bodyNum);
				$worksheet->write($a,73, "",$bodyPct);	
				
				$worksheet->write($a,74, $s->{C_RETAIL},$bodyNum);
				$worksheet->write($a,75, $s->{C_MARGIN},$bodyNum);
					if ($s->{C_RETAIL} le 0){
						$worksheet->write($a,76, "",$bodyPct); }
					else{
						$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$bodyPct); }
					
				$worksheet->write($a,77, "",$bodyNum);
				$worksheet->write($a,78, "",$bodyPct);
				$worksheet->write($a,79, "",$bodyNum);
				$worksheet->write($a,80, "",$bodyPct);	
				
				$worksheet->write($a,81, $s->{AMT502000},$bodyNum);
				$worksheet->write($a,82, "",$bodyNum);
				$worksheet->write($a,83, "",$bodyNum);
						
				$worksheet->write($a,84, $s->{AMT434000},$bodyNum);
				$worksheet->write($a,85, "",$bodyNum);
				$worksheet->write($a,86, "",$bodyNum);
						
				$worksheet->write($a,87, $s->{AMT458550},$bodyNum);
				$worksheet->write($a,88, "",$bodyNum);
				$worksheet->write($a,89, "",$bodyNum);
				
				$worksheet->write($a,90, $s->{AMT460100},$bodyNum);
				$worksheet->write($a,91, "",$bodyNum);
				$worksheet->write($a,92, "",$bodyNum);
					
				$worksheet->write($a,93, $s->{AMT460200},$bodyNum);
				$worksheet->write($a,94, "",$bodyNum);
				$worksheet->write($a,95, "",$bodyNum);
						
				$worksheet->write($a,96, $s->{AMT460300},$bodyNum);
				$worksheet->write($a,97, "",$bodyNum);
				$worksheet->write($a,98, "",$bodyNum);

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
			
			$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$headNumber);
			$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$headNumber);
				if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
					$worksheet->write($a,9, "",$headPct); }
				else{
					$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$headPct); }
				
			$worksheet->write($a,10, "",$headNumber);
			$worksheet->write($a,11, "",$headPct);
			$worksheet->write($a,12, "",$headNumber);
			$worksheet->write($a,13, "",$headPct);
			
			$worksheet->write($a,14, $s->{O_RETAIL},$headNumber);
			$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$headNumber);
				if ($s->{O_RETAIL} le 0){
					$worksheet->write($a,16, "",$headPct); }
				else{
					$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$headPct); }
				
			$worksheet->write($a,17, "",$headNumber);
			$worksheet->write($a,18, "",$headPct);
			$worksheet->write($a,19, "",$headNumber);
			$worksheet->write($a,20, "",$headPct);			
				
			$worksheet->write($a,21, $s->{O_RETAIL},$headNumber);
			$worksheet->write($a,22, $s->{O_MARGIN},$headNumber);
				if ($s->{O_RETAIL} le 0){
					$worksheet->write($a,23, "",$headPct); }
				else{
					$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$headPct); }
						
			$worksheet->write($a,24, "",$headNumber);
			$worksheet->write($a,25, "",$headPct);
			$worksheet->write($a,26, "",$headNumber);
			$worksheet->write($a,27, "",$headPct);	
				
			$worksheet->write($a,28, $s->{AMT501000},$headNumber);
			$worksheet->write($a,29, "",$headNumber);
			$worksheet->write($a,30, "",$headNumber);
				
			$worksheet->write($a,31, $s->{AMT503200},$headNumber);
			$worksheet->write($a,32, "",$headNumber);
			$worksheet->write($a,33, "",$headNumber);
				
			$worksheet->write($a,34, $s->{AMT503250},$headNumber);
			$worksheet->write($a,35, "",$headNumber);
			$worksheet->write($a,36, "",$headNumber);
				
			$worksheet->write($a,37, $s->{AMT503500},$headNumber);
			$worksheet->write($a,38, "",$headNumber);
			$worksheet->write($a,39, "",$headNumber);
				
			$worksheet->write($a,40, $s->{AMT506000},$headNumber);
			$worksheet->write($a,41, "",$headNumber);
			$worksheet->write($a,42, "",$headNumber);
			
			$worksheet->write($a,43, $s->{AMT503000},$headNumber);
			$worksheet->write($a,44, "",$headNumber);
			$worksheet->write($a,45, "",$headNumber);				
							
			$worksheet->write($a,46, $s->{AMT507000},$headNumber);
			$worksheet->write($a,47, "",$headNumber);
			$worksheet->write($a,48, "",$headNumber);
				
			$worksheet->write($a,49, $s->{AMT999998},$headNumber);
			$worksheet->write($a,50, "",$headNumber);
			$worksheet->write($a,51, "",$headNumber);
				
			$worksheet->write($a,52, $s->{AMT504000},$headNumber);
			$worksheet->write($a,53, "",$headNumber);
			$worksheet->write($a,54, "",$headNumber);
							
			$worksheet->write($a,55, $s->{AMT432000},$headNumber);
			$worksheet->write($a,56, "",$headNumber);
			$worksheet->write($a,57, "",$headNumber);
				
			$worksheet->write($a,58, $s->{AMT433000},$headNumber);
			$worksheet->write($a,59, "",$headNumber);
			$worksheet->write($a,60, "",$headNumber);
				
			$worksheet->write($a,61, $s->{AMT458490},$headNumber);
			$worksheet->write($a,62, "",$headNumber);
			$worksheet->write($a,63, "",$headNumber);
			
			$worksheet->write($a,64, $s->{AMT505000},$headNumber);
			$worksheet->write($a,65, "",$headNumber);
			$worksheet->write($a,66, "",$headNumber);
			
			$worksheet->write($a,67, $s->{C_RETAIL},$headNumber);
			$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$headNumber);
				if ($s->{C_RETAIL} le 0){
					$worksheet->write($a,69, "",$headPct); }
				else{
					$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$headPct); }						
			
			$worksheet->write($a,70, "",$headNumber);
			$worksheet->write($a,71, "",$headPct);
			$worksheet->write($a,72, "",$headNumber);
			$worksheet->write($a,73, "",$headPct);	
				
			$worksheet->write($a,74, $s->{C_RETAIL},$headNumber);
			$worksheet->write($a,75, $s->{C_MARGIN},$headNumber);
				if ($s->{C_RETAIL} le 0){
					$worksheet->write($a,76, "",$headPct); }
				else{
					$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$headPct); }
				
			$worksheet->write($a,77, "",$headNumber);
			$worksheet->write($a,78, "",$headPct);
			$worksheet->write($a,79, "",$headNumber);
			$worksheet->write($a,80, "",$headPct);	
				
			$worksheet->write($a,81, $s->{AMT502000},$headNumber);
			$worksheet->write($a,82, "",$headNumber);
			$worksheet->write($a,83, "",$headNumber);
						
			$worksheet->write($a,84, $s->{AMT434000},$headNumber);
			$worksheet->write($a,85, "",$headNumber);
			$worksheet->write($a,86, "",$headNumber);
					
			$worksheet->write($a,87, $s->{AMT458550},$headNumber);
			$worksheet->write($a,88, "",$headNumber);
			$worksheet->write($a,89, "",$headNumber);
				
			$worksheet->write($a,90, $s->{AMT460100},$headNumber);
			$worksheet->write($a,91, "",$headNumber);
			$worksheet->write($a,92, "",$headNumber);
					
			$worksheet->write($a,93, $s->{AMT460200},$headNumber);
			$worksheet->write($a,94, "",$headNumber);
			$worksheet->write($a,95, "",$headNumber);
						
			$worksheet->write($a,96, $s->{AMT460300},$headNumber);
			$worksheet->write($a,97, "",$headNumber);
			$worksheet->write($a,98, "",$headNumber);
			
			$a++;
		}

	$worksheet->write($a,7, $s->{O_RETAIL}+$s->{C_RETAIL},$headNumber);
	$worksheet->write($a,8, $s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$headNumber);
		if (($s->{O_RETAIL}+$s->{C_RETAIL}) le 0){
			$worksheet->write($a,9, "",$headPct); }
		else{
			$worksheet->write($a,9, ($s->{O_MARGIN}+$s->{C_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/($s->{O_RETAIL}+$s->{C_RETAIL}),$headPct); }
				
	$worksheet->write($a,10, "",$headNumber);
	$worksheet->write($a,11, "",$headPct);
	$worksheet->write($a,12, "",$headNumber);
	$worksheet->write($a,13, "",$headPct);
		
	$worksheet->write($a,14, $s->{O_RETAIL},$headNumber);
	$worksheet->write($a,15, $s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000},$headNumber);
		if ($s->{O_RETAIL} le 0){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($s->{O_MARGIN}+$s->{AMT501000}+$s->{AMT432000}+$s->{AMT433000}+$s->{AMT458490}+$s->{AMT503200}+$s->{AMT503250}+$s->{AMT503500}+$s->{AMT506000}+$s->{AMT503000}+$s->{AMT507000}+$s->{AMT999998}+$s->{AMT505000}+$s->{AMT504000})/$s->{O_RETAIL},$headPct); }
				
	$worksheet->write($a,17, "",$headNumber);
	$worksheet->write($a,18, "",$headPct);
	$worksheet->write($a,19, "",$headNumber);
	$worksheet->write($a,20, "",$headPct);			
	
	$worksheet->write($a,21, $s->{O_RETAIL},$headNumber);
	$worksheet->write($a,22, $s->{O_MARGIN},$headNumber);
		if ($s->{O_RETAIL} le 0){
			$worksheet->write($a,23, "",$headPct); }
		else{
			$worksheet->write($a,23, $s->{O_MARGIN}/$s->{O_RETAIL},$headPct); }
						
	$worksheet->write($a,24, "",$headNumber);
	$worksheet->write($a,25, "",$headPct);
	$worksheet->write($a,26, "",$headNumber);
	$worksheet->write($a,27, "",$headPct);	
				
	$worksheet->write($a,28, $s->{AMT501000},$headNumber);
	$worksheet->write($a,29, "",$headNumber);
	$worksheet->write($a,30, "",$headNumber);
				
	$worksheet->write($a,31, $s->{AMT503200},$headNumber);
	$worksheet->write($a,32, "",$headNumber);
	$worksheet->write($a,33, "",$headNumber);
				
	$worksheet->write($a,34, $s->{AMT503250},$headNumber);
	$worksheet->write($a,35, "",$headNumber);
	$worksheet->write($a,36, "",$headNumber);
				
	$worksheet->write($a,37, $s->{AMT503500},$headNumber);
	$worksheet->write($a,38, "",$headNumber);
	$worksheet->write($a,39, "",$headNumber);
			
	$worksheet->write($a,40, $s->{AMT506000},$headNumber);
	$worksheet->write($a,41, "",$headNumber);
	$worksheet->write($a,42, "",$headNumber);
				
	$worksheet->write($a,43, $s->{AMT503000},$headNumber);
	$worksheet->write($a,44, "",$headNumber);
	$worksheet->write($a,45, "",$headNumber);				
							
	$worksheet->write($a,46, $s->{AMT507000},$headNumber);
	$worksheet->write($a,47, "",$headNumber);
	$worksheet->write($a,48, "",$headNumber);
				
	$worksheet->write($a,49, $s->{AMT999998},$headNumber);
	$worksheet->write($a,50, "",$headNumber);
	$worksheet->write($a,51, "",$headNumber);
				
	$worksheet->write($a,52, $s->{AMT504000},$headNumber);
	$worksheet->write($a,53, "",$headNumber);
	$worksheet->write($a,54, "",$headNumber);
							
	$worksheet->write($a,55, $s->{AMT432000},$headNumber);
	$worksheet->write($a,56, "",$headNumber);
	$worksheet->write($a,57, "",$headNumber);
				
	$worksheet->write($a,58, $s->{AMT433000},$headNumber);
	$worksheet->write($a,59, "",$headNumber);
	$worksheet->write($a,60, "",$headNumber);
	
	$worksheet->write($a,61, $s->{AMT458490},$headNumber);
	$worksheet->write($a,62, "",$headNumber);
	$worksheet->write($a,63, "",$headNumber);
				
	$worksheet->write($a,64, $s->{AMT505000},$headNumber);
	$worksheet->write($a,65, "",$headNumber);
	$worksheet->write($a,66, "",$headNumber);
				
	$worksheet->write($a,67, $s->{C_RETAIL},$headNumber);
	$worksheet->write($a,68, $s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000},$headNumber);
		if ($s->{C_RETAIL} le 0){
			$worksheet->write($a,69, "",$headPct); }
		else{
			$worksheet->write($a,69, ($s->{C_MARGIN}+$s->{AMT434000}+$s->{AMT458550}+$s->{AMT460100}+$s->{AMT460200}+$s->{AMT460300}+$s->{AMT502000})/$s->{C_RETAIL},$headPct); }
						
	$worksheet->write($a,70, "",$headNumber);
	$worksheet->write($a,71, "",$headPct);
	$worksheet->write($a,72, "",$headNumber);
	$worksheet->write($a,73, "",$headPct);	
				
	$worksheet->write($a,74, $s->{C_RETAIL},$headNumber);
	$worksheet->write($a,75, $s->{C_MARGIN},$headNumber);
		if ($s->{C_RETAIL} le 0){
			$worksheet->write($a,76, "",$headPct); }
		else{
			$worksheet->write($a,76, $s->{C_MARGIN}/$s->{C_RETAIL},$headPct); }
		
	$worksheet->write($a,77, "",$headNumber);
	$worksheet->write($a,78, "",$headPct);
	$worksheet->write($a,79, "",$headNumber);
	$worksheet->write($a,80, "",$headPct);	
				
	$worksheet->write($a,81, $s->{AMT502000},$headNumber);
	$worksheet->write($a,82, "",$headNumber);
	$worksheet->write($a,83, "",$headNumber);
						
	$worksheet->write($a,84, $s->{AMT434000},$headNumber);
	$worksheet->write($a,85, "",$headNumber);
	$worksheet->write($a,86, "",$headNumber);
					
	$worksheet->write($a,87, $s->{AMT458550},$headNumber);
	$worksheet->write($a,88, "",$headNumber);
	$worksheet->write($a,89, "",$headNumber);
	
	$worksheet->write($a,90, $s->{AMT460100},$headNumber);
	$worksheet->write($a,91, "",$headNumber);
	$worksheet->write($a,92, "",$headNumber);
			
	$worksheet->write($a,93, $s->{AMT460200},$headNumber);
	$worksheet->write($a,94, "",$headNumber);
	$worksheet->write($a,95, "",$headNumber);
						
	$worksheet->write($a,96, $s->{AMT460300},$headNumber);
	$worksheet->write($a,97, "",$headNumber);
	$worksheet->write($a,98, "",$headNumber);

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
NVL((SUM((NVL(SALE_NET_VAL_OTR_TY,0))-(NVL(SALE_TOT_TAX_VAL_OTR_TY,0))-(NVL(SALE_TOT_DISC_VAL_OTR_TY,0))))/1000,0) NET_SALE_OTR_TY,
SUM(SALE_TOT_QTY_OTR_TY)/1000 SALE_TOT_QTY_OTR_TY, 
NVL((SUM((NVL(SALE_NET_VAL_CON_TY,0))-(NVL(SALE_TOT_TAX_VAL_CON_TY,0))-(NVL(SALE_TOT_DISC_VAL_CON_TY,0))))/1000,0) NET_SALE_CON_TY,
SUM(SALE_TOT_QTY_CON_TY)/1000 SALE_TOT_QTY_CON_TY, 
NVL((SUM((NVL(SALE_NET_VAL_OTR_LY,0))-(NVL(SALE_TOT_TAX_VAL_OTR_LY,0))-(NVL(SALE_TOT_DISC_VAL_OTR_LY,0))))/1000,0) NET_SALE_OTR_LY,
SUM(SALE_TOT_QTY_OTR_LY)/1000 SALE_TOT_QTY_OTR_LY, 
NVL((SUM((NVL(SALE_NET_VAL_CON_LY,0))-(NVL(SALE_TOT_TAX_VAL_CON_LY,0))-(NVL(SALE_TOT_DISC_VAL_CON_LY,0))))/1000,0) NET_SALE_CON_LY,
SUM(SALE_TOT_QTY_CON_LY)/1000 SALE_TOT_QTY_CON_LY
FROM (	
	SELECT TBL.PER, TBL.STORE_KEY STORE_KEY, TBL.DS_KEY DS_KEY, TBL.STORE_CODE STORE_CODE, 
		TY_OTR.SALE_NET_VAL_OTR_TY, TY_OTR.SALE_TOT_TAX_VAL_OTR_TY, TY_OTR.SALE_TOT_DISC_VAL_OTR_TY, TY_OTR.SALE_TOT_QTY_OTR_TY,
		TY_CON.SALE_NET_VAL_CON_TY, TY_CON.SALE_TOT_TAX_VAL_CON_TY, TY_CON.SALE_TOT_DISC_VAL_CON_TY, TY_CON.SALE_TOT_QTY_CON_TY, 
		LY_OTR.SALE_NET_VAL_OTR_LY, LY_OTR.SALE_TOT_TAX_VAL_OTR_LY, LY_OTR.SALE_TOT_DISC_VAL_OTR_LY, LY_OTR.SALE_TOT_QTY_OTR_LY,
		LY_CON.SALE_NET_VAL_CON_LY, LY_CON.SALE_TOT_TAX_VAL_CON_LY, LY_CON.SALE_TOT_DISC_VAL_CON_LY, LY_CON.SALE_TOT_QTY_CON_LY,
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
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_OTR_TY, 
		SUM(SALE_TOT_DISC_VAL) SALE_TOT_DISC_VAL_OTR_TY
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
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_CON_TY, 
		SUM(SALE_TOT_DISC_VAL) SALE_TOT_DISC_VAL_CON_TY
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
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_OTR_LY, 
		SUM(SALE_TOT_DISC_VAL) SALE_TOT_DISC_VAL_OTR_LY
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
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_CON_LY, 
		SUM(SALE_TOT_DISC_VAL) SALE_TOT_DISC_VAL_CON_LY
		FROM AGG_DLY_STR_PROD AGG 
			INNER JOIN DIM_PRODUCT M ON AGG.PRODUCT_KEY = M.PRODUCT_KEY AND M.PRODUCT_CATEGORY = 2 
			INNER JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
			INNER JOIN DIM_DATE DA ON AGG.DATE_KEY = DA.DATE_KEY
		WHERE AGG.DATE_KEY BETWEEN $wk_st_date_key_ly AND $wk_en_date_key_ly 
		GROUP BY TO_CHAR(DA.DATE_FLD, 'MON'), STORE_KEY, DS_KEY, STORE_CODE)LY_CON
		
		ON TBL.PER = LY_CON.PER AND TBL.STORE_KEY = LY_CON.STORE_KEY AND TBL.STORE_CODE = LY_CON.STORE_CODE AND TBL.DS_KEY = LY_CON.DS_KEY
	
	UNION ALL 
	
	SELECT TO_CHAR(DP.DATE_FLD, 'MON') PER, STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE, 
		0 AS SALE_NET_VAL_OTR_TY, 0 AS SALE_TOT_TAX_VAL_OTR_TY, 0 AS SALE_TOT_DISC_VAL_OTR_TY, 0 AS SALE_TOT_QTY_OTR_TY,
		0 AS SALE_NET_VAL_OTR_LY, 0 AS SALE_TOT_TAX_VAL_OTR_LY, 0 AS SALE_TOT_DISC_VAL_OTR_LY, 0 AS SALE_TOT_QTY_CON_TY, 
		0 AS SALE_NET_VAL_CON_TY, 0 AS SALE_TOT_TAX_VAL_CON_TY, 0 AS SALE_TOT_DISC_VAL_CON_LY, 0 AS SALE_TOT_QTY_OTR_LY,
		0 AS SALE_NET_VAL_CON_LY, 0 AS SALE_TOT_TAX_VAL_CON_LY, 0 AS SALE_TOT_DISC_VAL_CON_LY, 0 AS SALE_TOT_QTY_CON_LY,		
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

my( $to, $subject, $msgbody_file, $attachment_file_1, $attachment_file_2 ) = @ARGV;

$to = ' frank.gaisano@metrogaisano.com, arthur.emmanuel@metrogaisano.com, gerry.guanlao@metrogaisano.com, eric.redona@metrogaisano.com, lucille.malazarte@metrogaisano.com, melissa.catan@metrogaisano.com, rex.cabanilla@metrogaisano.com, karan.malani@metrogaisano.com, lia.chipeco@metrogaisano.com, chit.lazaro@metrogaisano.com, fili.mercado@metrogaisano.com, margaret.ang@metrogaisano.com, luz.bitang@metrogaisano.com, emily.silverio@metrogaisano.com ';

$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com ';
		
# $to = ' kent.mamalias@metrogaisano.com ';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'Total Margin Performance ' . $as_of;

$msgbody_file = 'message_margin.txt';

$attachment_file_1 = "TOTAL MARGIN PERFORMANCE (as of $as_of) v1.6.xlsx";
$attachment_file_2 = "TOTAL MARGIN PERFORMANCE (as of $as_of) v1.6.pdf";

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
Content-Type: text/html; charset="iso-8859-1"
Content-Transfer-Encoding: quoted-printable

<html>
Hi All, <br> <br>
Please see attached Total Margin Performance Report. <br> <br>

If you need assistance, kindly email arcbi.support&#64;metrogaisano.com. <br> <br>

<b>Disclaimer</b><p>
<table>
The figures in this report are tentative and serve to show a quick determination of your margins on a weekly basis.

<p>This report shall be generated every Monday.  Since the month-end rebates and cost allocations are not available weekly, consider the figures as tentative.  The final and official margins shall be reported by Finance at the end of every month.

<p>Tentative margins for Fresh - Meat, Seafood, Produce, Suisse Cottage/Bakeshop, Metro Gourmet/Cafe, Food Avenue are based on perpetual method of inventory accounting whereas the final margins shall be based on the periodic method. 
</table><br>

Regards, <br>
ARC BI Support <p>
</html>

$boundary
Content-Type: application/octet-stream; name="$attachment_file_1"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$attachment_file_1"
$attachment_data_1
Content-Type: application/octet-stream; name="$attachment_file_2"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$attachment_file_2"
$attachment_data_2
$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

}

sub mail_grp1_external {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1, $attachment_file_2 ) = @ARGV;

$to = ' artemm12@aol.com, frankgaisano@gmail.com ';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'Total Margin Performance ' . $as_of;

$msgbody_file = 'message_margin.txt';

$attachment_file_1 = "TOTAL MARGIN PERFORMANCE (as of $as_of) v1.6.xlsx";
$attachment_file_2 = "TOTAL MARGIN PERFORMANCE (as of $as_of) v1.6.pdf";

my $msgbody = read_file( $msgbody_file );

my $attachment_data_1 = encode_base64( read_file( $attachment_file_1, 1 ));
my $attachment_data_2 = encode_base64( read_file( $attachment_file_2, 1 ));

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
Content-Transfer-Encoding: quoted-printable

<html>
Hi All, <br> <br>
Please see attached Total Margin Performance Report. <br> <br>

If you need assistance, kindly email arcbi.support&#64;metrogaisano.com. <br> <br>

<b>Disclaimer</b><p>
<table>
The figures in this report are tentative and serve to show a quick determination of your margins on a weekly basis.

<p>This report shall be generated every Monday.  Since the month-end rebates and cost allocations are not available weekly, consider the figures as tentative.  The final and official margins shall be reported by Finance at the end of every month.

<p>Tentative margins for Fresh - Meat, Seafood, Produce, Suisse Cottage/Bakeshop, Metro Gourmet/Cafe, Food Avenue are based on perpetual method of inventory accounting whereas the final margins shall be based on the periodic method. 
</table><br>

Regards, <br>
ARC BI Support <p>
</html>

$boundary
Content-Type: application/octet-stream; name="$attachment_file_1"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$attachment_file_1"
$attachment_data_1
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







