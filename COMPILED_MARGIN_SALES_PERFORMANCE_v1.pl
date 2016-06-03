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


$test_query = qq{ SELECT CASE WHEN EXISTS (SELECT SEQ_NO, ETL_SUMMARY, VALUE, ARC_DATE FROM ADMIN_ETL_SUMMARY WHERE TO_DATE(ARC_DATE, 'DD-MON-YY') = TO_DATE(SYSDATE,'DD-MON-YY')) THEN 1 ELSE 0 END STATUS FROM DUAL };

# $test_query = qq{ SELECT CASE WHEN EXISTS (SELECT *
					# FROM ADMIN_ETL_LOG 
					# WHERE TO_DATE(LOG_DATE, 'DD-MON-YY') = TO_DATE(SYSDATE,'DD-MON-YY') AND TASK_ID = 'AggDlyStrProd' AND ERR_CODE = 0) THEN 1 ELSE 0 END STATUS 
					# FROM DUAL };

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
	
	#=============================== GROUP 1=================================================# 
 
	$workbook = Excel::Writer::XLSX->new("RMS_LY_CONSOLIDATED MARGIN - Summary (as of $as_of) v1.xlsx");
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
	
	&generate_csv;
	
	&new_sheet($sheet = "Summary");
	&call_str;

	# &new_sheet($sheet = "GenMerch_Spmkt");
	# &call_str_merchandise;
	
	&new_sheet_2($sheet = "Department");			
	&call_div;
		
	$workbook->close();
	
	# my $pdf_job_1 = Win32::Job->new;
	# $pdf_job_1->spawn( "cmd" , q{cmd /C java ecp_FileConverter "Daily Sales Performance - Summary (as of } . $as_of . q{) v1.3.xlsx" pdf});
	# $pdf_job_1->run(60);	
	
	# &mail_grp1;	
	
	##=============================== GROUP 2=================================================##

	# $workbook = Excel::Writer::XLSX->new("Daily Sales Performance (as of $as_of).xlsx");
	# $bold = $workbook->add_format( bold => 1, size => 14 );
	# $bold1 = $workbook->add_format( bold => 1, size => 16 );
	# $script = $workbook->add_format( size => 8, italic => 1 );
	# $bold2 = $workbook->add_format( size => 11 );
	# $border1 = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3 );
	# $border2 = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', rotation => 90, text_wrap =>1, size => 10, shrink => 1 );
	# $code = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10 );
	# $desc = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10 );
	# $ponkan = $workbook->set_custom_color( 53, 254, 238, 230);
	# $abo = $workbook->set_custom_color( 16, 220, 218, 219);
	# $sky = $workbook->set_custom_color( 12, 205, 225, 255);
	# $pula = $workbook->set_custom_color( 10, 255, 189, 189);
	# $lumot = $workbook->set_custom_color( 17, 196, 189, 151);
	# $comp = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10,  bg_color => $lumot, bold => 1 );
	# $all = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10,  bg_color => $abo, bold => 1 );
	# $headN = $workbook->add_format( border => 1, align => 'center', valign => 'center', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
	# $headD = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 9, bg_color => $abo, bold => 1 );
	# $headPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
	# $headNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3, bg_color => $abo, bold => 1 );
	# $headNumber = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 3, bg_color => $abo, bold => 1 );
	# $head = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
	# $subhead = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => 9, bg_color => $ponkan, bold => 1 );
	# $bodyN = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
	# $bodyPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
	# $bodyNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 3,  bold => 1);
	# $body = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
	# $subt = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 9);
	# $down = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => 9, bg_color => $pula );

	# printf "Arc BI Sales Performance Part 2 \n";

	# &new_sheet($sheet = "Summary");
	# &call_str;

	# $workbook->close();
	
	# my $pdf_job_2 = Win32::Job->new;
	# $pdf_job_2->spawn( "cmd" , q{cmd /C java ecp_FileConverter "Daily Sales Performance (as of } . $as_of . q{).xlsx" pdf});
	# $pdf_job_2->run(60);	

	# &mail_grp2;

	$tst_query->finish();
	$dbh_csv->disconnect;
	$dbh->disconnect; 
	
	exit;
	
}

else{
	print "Test Status = ". $test ." \nETL still running\n";
	
	$tst_query->finish();
	
	sleep(600);
	
	goto START;
}
 
#================================= FUNCTIONS ==================================#

sub call_div {

$a = 10, $e = 10, $counter = 0;
$grp_o_retail = 0, $grp_o_margin = 0, $grp_c_retail = 0, $grp_c_margin = 0;
$total_o_retail = 0, $total_o_margin = 0, $total_c_retail = 0, $total_c_margin = 0;
$type_test = 0;

$worksheet->write($a-9, 3, "Front Margin Performance", $bold1);
#$worksheet->write($a-8, 3, "MTD: $mo_st_date_fld - $mo_en_date_fld");
$worksheet->write($a-8, 3, "MTD: 01 Oct 2013 - 31 Oct 2013");
$worksheet->write($a-7, 3, "As of $as_of");

##========================= COMP STORES ===========================##

&heading_2;
&heading;
&query_dept($new_flg = 0, $matured_flg = 1, $loc_desc = "COMP STORES");

##========================= ALL STORES ===========================##

$a += 7;
$grp_o_retail = 0, $grp_o_margin = 0, $grp_c_retail = 0, $grp_c_margin = 0;
$total_o_retail = 0, $total_o_margin = 0, $total_c_retail = 0, $total_c_margin = 0;
$type_test = 0;

&heading_2;
&heading;
&query_dept($new_flg = 1, $matured_flg = 1, $loc_desc = "ALL STORES");

##========================= BY STORE ===========================##

# foreach my $i ( '2001', '2001W', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', '2013', '2223', '3000', '3007', '3008', '3009', '3010', '3011', '3013', '4003', '4004', '6001', '6002', '6003', '6004', '6005', '6006', '6009', '6010', '6011', '6012', '6013' ){ 

	# $a += 7;
	# $total_o_retail = 0, $total_o_margin = 0, $total_c_retail = 0, $total_c_margin = 0;
	# $grp_o_retail = 0, $grp_o_margin = 0, $grp_c_retail = 0, $grp_c_margin = 0;
	# &heading_2;
	# &heading;
	# &query_dept_store($store = $i);

# }

}

sub call_str {

$a=10, $d=0, $e=10, $counter=0, $comp_su = 0, $new_su = 0, $comp_nb = 0, $comp_hy = 0, $new_hy = 0, $comp_ds = 0, $new_ds = 0, $type_test = 1;

$worksheet->write($a-9, 3, "Front Margin Performance", $bold1);
#$worksheet->write($a-8, 3, "MTD: $mo_st_date_fld - $mo_en_date_fld");
$worksheet->write($a-8, 3, "MTD: 01 Oct 2013 - 31 Oct 2013");
$worksheet->write($a-7, 3, "As of $as_of");

$worksheet->write($a-4, 3, "Summary", $bold);

&heading;

$worksheet->merge_range( $a-2, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a-1, 6, 'Format', $subhead );

&strComp_Ds;
&strComp_Su;
&strComp_Hy;
&strComp_Nb;

&strNew_Ds;
&strNew_Su;
&strNew_Hy;
&strNew_Nb;

&str_Ds;
&str_Su;
&str_Hy;
&str_Nb;

$type_test = 2;

$worksheet->write($a-4, 3, "Per Store", $bold);

&heading;

$worksheet->merge_range( $a-2, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a-1, 4, 'Type', $subhead );
$worksheet->merge_range( $a-2, 5, $a-1, 5, 'Code', $subhead );
$worksheet->merge_range( $a-2, 6, $a-1, 6, 'Desc', $subhead );

&strComp_Ds;
&strNew_Ds;
&strComp_Su;
&strNew_Su;
&strComp_Hy;
&strNew_Hy;
&strComp_Nb;

}

sub call_str_merchandise {

$a=10, $d=0, $e=10, $counter=0, $comp_su = 0, $new_su = 0, $comp_nb = 0, $comp_hy = 0, $new_hy = 0, $comp_ds = 0, $new_ds = 0, $type_test = 3;

$worksheet->write($a-9, 3, "Front Margin Performance", $bold1);
#$worksheet->write($a-8, 3, "MTD: $mo_st_date_fld - $mo_en_date_fld");
$worksheet->write($a-8, 3, "MTD: 01 Oct 2014 - 31 Oct 2014");
$worksheet->write($a-7, 3, "As of $as_of");

$worksheet->write($a-4, 3, "Summary", $bold);

&heading_3;

$worksheet->merge_range( $a-2, 3, $a, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a, 6, 'Format', $subhead );

$a += 1;

&strComp_Ds;
&strComp_Su;
&strComp_Hy;
&strComp_Nb;

&strNew_Ds;
&strNew_Su;
&strNew_Hy;
&strNew_Nb;

&str_Ds;
&str_Su;
&str_Hy;
&str_Nb;

$type_test = 4;

$worksheet->write($a-4, 3, "Per Store", $bold);

&heading_3;

$worksheet->merge_range( $a-2, 3, $a, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a, 4, 'Type', $subhead );
$worksheet->merge_range( $a-2, 5, $a, 5, 'Code', $subhead );
$worksheet->merge_range( $a-2, 6, $a, 6, 'Desc', $subhead );

$a += 1;

&strComp_Ds;
&strNew_Ds;
&strComp_Su;
&strNew_Su;
&strComp_Hy;
&strNew_Hy;
&strComp_Nb;

}


sub strComp_Su {

$div_name = "Comp";  $div_name3 = "Supermarket";
$mrch1 = 'SU'; $mrch2 = 'SU'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8040'; $dept_grp2 = '0000'; $dept_grp3 = '0000'; $dept_grp4 = '0000'; $dept_grp5 = '0000'; $dept_grp6 = '0000'; $dept_grp7 = '0000';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_counter = 0, $s1f2_row = 0;

$store1 = 'SU2001'; $store2 = 'SU2002'; $store3 = 'SU2003'; $store4 = 'SU2004'; $store5 = 'SU2006'; $store6 = 'SU2007'; $store7 = 'SU2009'; $store8 = 'SU2012'; $store9 = 'SU2001W'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2012'; $stor9 = '2001W'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 

	elsif($type_test eq 2){	
		
		&query_by_store;
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );

		$tst=$a-$counter; $comp_su=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER
	}
	
	elsif($type_test eq 3){	&query_summary_merchandise;	} 
	
	elsif($type_test eq 4){	
		
		&query_by_store_merchandise;
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );

		$tst=$a-$counter; $comp_su=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER
	}
	
}

sub strNew_Su {

$div_name = "New"; $div_name2 = "Supermarket";  $div_name3 = "Supermarket";
$mrch1 = 'SU'; $mrch2 = 'SU'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8040'; $dept_grp2 = '0000'; $dept_grp3 = '0000'; $dept_grp4 = '0000'; $dept_grp5 = '0000'; $dept_grp6 = '0000'; $dept_grp7 = '0000';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_counter = 0, $s1f2_row = 0;

$store1 = 'SU2013'; $store2 = 'SU4004'; $store3 = 'SU3009'; $store4 = 'SU3010'; $store5 = 'SU3011'; $store6 = '0000'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2013'; $stor2 = '4004'; $stor3 = '3009'; $stor4 = '3010'; $stor5 = '3011'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 

	elsif($type_test eq 2) {	
		
		&query_by_store;
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_su=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 14, 15, 17, 19, 21, 22, 24, 26 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_su,$col).','.xl_rowcol_to_cell($new_su,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} }

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
		
	}

	elsif($type_test eq 3){	&query_summary_merchandise;	} 
	
	elsif($type_test eq 4) {	
		
		&query_by_store_merchandise;
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_su=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 13, 15, 17, 18, 20, 22, 23, 25, 27, 28, 30, 32, 33, 35 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_su,$col).','.xl_rowcol_to_cell($new_su,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 13 or $col eq 18 or $col eq 23 or $col eq 28 or $col eq 33){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct );
						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
						}
					}
					
					elsif ($col eq 10 or $col eq 15 or $col eq 20 or $col eq 25 or $col eq 30 or $col eq 35){
						if( xl_rowcol_to_cell( $a, $col-3 ) le 0){
							$worksheet->write( $a, $col+1, '0', $bodyPct );
						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
						}
					}
			}

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
		
	}
}

sub strComp_Nb {

$div_name = "Comp"; $div_name2 = "Neighborhood";  $div_name3 = "Neighborhood Store";
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000';  $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_counter = 0, $s1f2_row = 0;

$store1 = 'SU3001'; $store2 = 'SU3007'; $store3 = 'SU4003'; $store4 = 'DS3001'; $store5 = 'DS3007'; $store6 = 'DS4003'; $store7 = 'OT3001'; $store8 = 'OT3007'; $store9 = 'OT4003'; $store10 = 'SU3002'; $store11 = 'DS3002'; $store12 = 'OT3002'; $store13 = 'SU3003'; $store14 = 'DS3003'; $store15 = 'OT3003'; $store16 = 'SU3004'; $store17 = 'DS3004'; $store18 = 'OT3004'; $store19 = 'SU3005'; $store20 = 'DS3005'; $store21 = 'OT3005'; $store22 = 'SU3006'; $store23 = 'DS3006'; $store24 = 'OT3006'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '30001'; $stor2 = '3007'; $stor3 = '4003'; $stor4 = '3002'; $stor5 = '3003'; $stor6 = '3004'; $stor7 = '3005'; $stor8 = '3006'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000'; 

	if($type_test eq 1){	
	
		&query_summary;	
		
		$counter = 4; 
		&calc8; 
		
		$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
		$worksheet->merge_range( $a-$counter, 3, $a, 3, $div_name, $border2 );
		$a+=1; $counter = 0; $d=$a;
		
	} 
	
	elsif($type_test eq 2) {	

		&query_by_store;	
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $a-$counter, 3, $a+1, 3, $div_name2, $border2 );

		$tst = $a; $comp_nb=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 14, 15, 17, 19, 21, 22, 24, 26 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} }

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
	
	}
	
	elsif($type_test eq 3){	
	
		&query_summary_merchandise;	
		
		$counter = 4; 
		&calc8; 
		
		$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
		$worksheet->merge_range( $a-$counter, 3, $a, 3, $div_name, $border2 );
		$a+=1; $counter = 0; $d=$a;
		
	} 	
	
	elsif($type_test eq 4) {	

		&query_by_store_merchandise;	
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $a-$counter, 3, $a+1, 3, $div_name2, $border2 );

		$tst = $a; $comp_nb=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 13, 15, 17, 18, 20, 22, 23, 25, 27, 28, 30, 32, 33, 35 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 13 or $col eq 18 or $col eq 23 or $col eq 28 or $col eq 33){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct ); 						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct ); 						}
					}
					
					elsif ($col eq 10 or $col eq 15 or $col eq 20 or $col eq 25 or $col eq 30 or $col eq 35){
						if( xl_rowcol_to_cell( $a, $col-3 ) le 0){
							$worksheet->write( $a, $col+1, '0', $bodyPct ); 						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct ); 						}
					}
			}

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
	
	}
	
}

sub strNew_Nb {

$div_name = "New"; $div_name2 = "Neighborhood";  $div_name3 = "Neighborhood Store";
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '0000'; $division_grp2 = '0000';  $division_grp3 = '0000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_counter = 0, $s1f2_row = 0;

$store1 = '0000'; $store2 = '0000'; $store3 = '0000'; $store4 = '0000'; $store5 = '0000'; $store6 = '0000'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '0000'; $stor2 = '0000'; $stor3 = '0000'; $stor4 = '0000'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000';  

	if($type_test eq 1){	
	
		&query_summary;	
		
		$counter = 4; 
		&calc8; 
		
		$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
		$worksheet->merge_range( $a-$counter, 3, $a, 3, $div_name, $border2 );
		$a+=1; $counter = 0; $d=$a;
		
	} 
	
	elsif($type_test eq 2) {	

		&query_by_store;	
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $a-$counter, 3, $a+1, 3, $div_name2, $border2 );

		$tst = $a; $comp_nb=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 14, 15, 17, 19, 21, 22, 24, 26 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} }

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
	
	}
	
	if($type_test eq 3){	
	
		&query_summary_merchandise;	
		
		$counter = 4; 
		&calc8; 
		
		$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
		$worksheet->merge_range( $a-$counter, 3, $a, 3, $div_name, $border2 );
		$a+=1; $counter = 0; $d=$a;
		
	} 

}

sub strComp_Hy {

$div_name = "Comp";  $div_name3 = "Hypermarket";
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_counter = 0, $s1f2_row = 0;

$store1 = 'SU2005'; $store2 = 'SU2008'; $store3 = 'SU2010'; $store4 = 'SU2011'; $store5 = 'DS2005'; $store6 = 'DS2008'; $store7 = 'DS2010'; $store8 = 'DS2011'; $store9 = 'DS2005'; $store10 = 'OT2005'; $store11 = 'OT2008'; $store12 = 'OT2010'; $store13 = 'OT2011'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2005'; $stor2 = '2008'; $stor3 = '2010'; $stor4 = '2011'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000';  

	if($type_test eq 1){	&query_summary;	} 

	elsif($type_test eq 2) {	

		&query_by_store;	
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );

		$comp_hy=$a; $tst = $a-$counter; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER
	}
	
	elsif($type_test eq 3){	&query_summary_merchandise;	} 
	
	elsif($type_test eq 4) {	

		&query_by_store_merchandise;	
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );

		$comp_hy=$a; $tst = $a-$counter; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER
	}
}

sub strNew_Hy {

$div_name = "New"; $div_name2 = "Hypermarket";  $div_name3 = "Hypermarket";
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_counter = 0, $s1f2_row = 0;

$store1 = 'SU6001'; $store2 = 'SU6002'; $store3 = 'SU6003'; $store4 = 'SU6004'; $store5 = 'SU6005'; $store6 = 'SU6012'; $store7 = 'SU6009'; $store8 = 'SU6010'; $store9 = 'SU6011'; $store10 = 'DS6001'; $store11 = 'DS6002'; $store12 = 'DS6003'; $store13 = 'DS6004'; $store14 = 'DS6005'; $store15 = 'DS6012'; $store16 = 'DS6009'; $store17 = 'DS6010'; $store18 = 'DS6011'; $store19 = 'OT6002'; $store20 = 'OT6003'; $store21 = 'OT6004'; $store22 = 'OT6005'; $store23 = 'OT6012'; $store24 = 'OT6009'; $store25 = 'OT6010'; $store26 = 'OT6011';  $store27 = 'OT6000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '6001'; $stor2 = '6001'; $stor3 = '6003'; $stor4 = '6004'; $stor5 = '6005'; $stor6 = '6012'; $stor7 = '6009'; $stor8 = '6010'; $stor9 = '6011'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000';  

	if($type_test eq 1){	
	
		&query_summary;	
		$counter = 4; 
		$counter = 0; $d=$a; 
		
	} 

	elsif($type_test eq 2) {	

		&query_by_store;	
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_hy=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 14, 15, 17, 19, 21, 22, 24, 26 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_hy,$col).','.xl_rowcol_to_cell($new_hy,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} }

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
	
	}
	
	if($type_test eq 3){	
	
		&query_summary_merchandise;	
		$counter = 4; 
		$counter = 0; $d=$a; 
		
	}

	elsif($type_test eq 4) {	

		&query_by_store_merchandise;	
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_hy=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 13, 15, 17, 18, 20, 22, 23, 25, 27, 28, 30, 32, 33, 35 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_hy,$col).','.xl_rowcol_to_cell($new_hy,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 13 or $col eq 18 or $col eq 23 or $col eq 28 or $col eq 33){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct ); 						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct ); 						}
					}
					
					elsif ($col eq 10 or $col eq 15 or $col eq 20 or $col eq 25 or $col eq 30 or $col eq 35){
						if( xl_rowcol_to_cell( $a, $col-3 ) le 0){
							$worksheet->write( $a, $col+1, '0', $bodyPct ); 						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct ); 						}
					}
			}

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
	
	}
	
}

sub strComp_Ds {

$div_name = "Comp";  $div_name3 = "Department Store";
$mrch1 = 'DS'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '9000'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8050'; $dept_grp5 = '8060'; $dept_grp6 = '8070'; $dept_grp7 = '0000';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_counter = 0, $s1f2_row = 0;

$store1 = 'DS2001'; $store2 = 'DS2002'; $store3 = 'DS2003'; $store4 = 'DS2004'; $store5 = 'DS2006'; $store6 = 'DS2007'; $store7 = 'DS2009'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 

	elsif($type_test eq 2) {	
		
		&query_by_store;	
		&calc8;
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		
		$comp_ds=$a; $tst = $a-$counter; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER
	}
	
	elsif($type_test eq 3){	&query_summary_merchandise;	} 
	
	elsif($type_test eq 4) {	
		
		&query_by_store_merchandise;	
		&calc8;
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		
		$comp_ds=$a; $tst = $a-$counter; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER
	}

}

sub strNew_Ds {

$div_name = "New"; $div_name2 = "Department Store"; $div_name3 = "Department Store";
$mrch1 = 'DS'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '9000'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8050'; $dept_grp5 = '8060'; $dept_grp6 = '8070'; $dept_grp7 = '0000';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_counter = 0, $s1f2_row = 0;

$store1 = 'DS2223'; $store2 = '0000'; $store3 = '0000'; $store4 = '0000'; $store5 = '0000'; $store6 = '0000'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2223'; $stor2 = '0000'; $stor3 = '0000'; $stor4 = '0000'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 

	elsif($type_test eq 2) {	
		
		&query_by_store;	
		&calc8;
		
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_ds=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 14, 15, 17, 19, 21, 22, 24, 26 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_ds,$col).','.xl_rowcol_to_cell($new_ds,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26){
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} }

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
		
	}
		
	elsif($type_test eq 3){	&query_summary_merchandise;	}	
	
	elsif($type_test eq 4) {	
		
		&query_by_store_merchandise;	
		&calc8;
		
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_ds=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 13, 15, 17, 18, 20, 22, 23, 25, 27, 28, 30, 32, 33, 35 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_ds,$col).','.xl_rowcol_to_cell($new_ds,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 13 or $col eq 18 or $col eq 23 or $col eq 28 or $col eq 33){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct ); 						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct ); 						}
					}
					
					elsif ($col eq 10 or $col eq 15 or $col eq 20 or $col eq 25 or $col eq 30 or $col eq 35){
						if( xl_rowcol_to_cell( $a, $col-3 ) le 0){
							$worksheet->write( $a, $col+1, '0', $bodyPct ); 						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct ); 						}
					}
			}

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
		
	}
	
}


sub str_Ds {

$div_name = "Comp"; $div_name3 = 'Department Store';
$mrch1 = 'DS'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '9000'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8050'; $dept_grp5 = '8060'; $dept_grp6 = '8070'; $dept_grp7 = '0000';
$s1f2_wtd_net_ty = 0, $s1f2_wtd_net_ly = 0, $s1f2_wtd_target = 0;
$s1f2_mtd_net_ty = 0,	$s1f2_mtd_net_ly = 0, $s1f2_mtd_target = 0;
$s1f2_qtd_net_ty = 0,	$s1f2_qtd_net_ly = 0, $s1f2_qtd_target = 0;
$s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;

$store1 = 'DS2001'; $store2 = 'DS2002'; $store3 = 'DS2003'; $store4 = 'DS2004'; $store5 = 'DS2006'; $store6 = 'DS2007'; $store7 = 'DS2009'; $store8 = 'DS2223'; $store9 = '0000'; $store10 = 'OOOO'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2223'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 
	elsif($type_test eq 3){	&query_summary_merchandise;	} 

}

sub str_Su {

$div_name = "Comp"; $div_name3 = 'Supermarket';
$mrch1 = 'SU'; $mrch2 = 'SU'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8040'; $dept_grp2 = '0000'; $dept_grp3 = '0000'; $dept_grp4 = '0000'; $dept_grp5 = '0000'; $dept_grp6 = '0000'; $dept_grp7 = '0000';
$s1f2_wtd_net_ty = 0, $s1f2_wtd_net_ly = 0, $s1f2_wtd_target = 0;
$s1f2_mtd_net_ty = 0,	$s1f2_mtd_net_ly = 0, $s1f2_mtd_target = 0;
$s1f2_qtd_net_ty = 0,	$s1f2_qtd_net_ly = 0, $s1f2_qtd_target = 0;
$s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;

$store1 = 'SU2001'; $store2 = 'SU2002'; $store3 = 'SU2003'; $store4 = 'SU2004'; $store5 = 'SU2001W'; $store6 = 'SU2006'; $store7 = 'SU2007'; $store8 = '0000'; $store9 = 'SU2009'; $store10 = 'SU2013'; $store11 = 'SU4004'; $store12 = 'SU2012'; $store13 = 'SU3009'; $store14 = 'SU3010'; $store15 = 'SU3011'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2013'; $stor9 = '2012'; $stor10 = '4004'; $stor11 = '3009'; $stor12 = '3010';  $stor13 = '3011';    $stor14 = '2001W'; 

	if($type_test eq 1){	&query_summary;	} 
	elsif($type_test eq 3){	&query_summary_merchandise;	} 
		
}

sub str_Hy {

$div_name = "New"; $div_name3 = 'Hypermarket';
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_wtd_net_ty = 0, $s1f2_wtd_net_ly = 0, $s1f2_wtd_target = 0;
$s1f2_mtd_net_ty = 0,	$s1f2_mtd_net_ly = 0, $s1f2_mtd_target = 0;
$s1f2_qtd_net_ty = 0,	$s1f2_qtd_net_ly = 0, $s1f2_qtd_target = 0;
$s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU6001'; $store2 = 'SU6002'; $store3 = 'SU6003'; $store4 = 'SU6004'; $store5 = 'SU6005'; $store6 = 'SU6012'; $store7 = 'SU6009'; $store8 = 'SU6010'; $store9 = 'SU6011'; $store10 = 'DS6001'; $store11 = 'DS6002'; $store12 = 'DS6003'; $store13 = 'DS6004'; $store14 = 'DS6005'; $store15 = 'DS6012'; $store16 = 'DS6009'; $store17 = 'DS6010'; $store18 = 'DS6011'; $store19 = 'OT6002'; $store20 = 'OT6003'; $store21 = 'OT6004'; $store22 = 'OT6005'; $store23 = 'OT6012'; $store24 = 'OT6009'; $store25 = 'OT6010'; $store26 = 'OT6011';  $store27 = 'OT6000'; $store28 = 'SU2005'; $store29 = 'SU2008'; $store30 = 'SU2010'; $store31 = 'SU2011'; $store32 = 'DS2005'; $store33 = 'DS2008'; $store34 = 'DS2010'; $store35 = 'DS2011'; $store36 = 'DS2005'; $store37 = 'OT2005'; $store38 = 'OT2008'; $store39 = 'OT2010'; $store40 = 'OT2011';

$stor1 = '6001'; $stor2 = '6002'; $stor3 = '6003'; $stor4 = '6004'; $stor5 = '6005'; $stor6 = '6012'; $stor7 = '6009'; $stor8 = '6010'; $stor9 = '6011'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 
	elsif($type_test eq 3){	&query_summary_merchandise;	} 

}

sub str_Nb {

$div_name = "All"; $div_name3 = 'Neighborhood Store';
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000';  $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_wtd_net_ty = 0, $s1f2_wtd_net_ly = 0, $s1f2_wtd_target = 0;
$s1f2_mtd_net_ty = 0,	$s1f2_mtd_net_ly = 0, $s1f2_mtd_target = 0;
$s1f2_qtd_net_ty = 0,	$s1f2_qtd_net_ly = 0, $s1f2_qtd_target = 0;
$s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU3001'; $store2 = 'SU3007'; $store3 = 'SU4003'; $store4 = 'DS3001'; $store5 = 'DS3007'; $store6 = 'DS4003'; $store7 = 'OT3001'; $store8 = 'OT3007'; $store9 = 'OT4003'; $store10 = 'SU3002'; $store11 = 'DS3002'; $store12 = 'OT3002'; $store13 = 'SU3003'; $store14 = 'DS3003'; $store15 = 'OT3003'; $store16 = 'SU3004'; $store17 = 'DS3004'; $store18 = 'OT3004'; $store19 = 'SU3005'; $store20 = 'DS3005'; $store21 = 'OT3005'; $store22 = 'SU3006'; $store23 = 'DS3006'; $store24 = 'OT3006'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '30001'; $stor2 = '3007'; $stor3 = '4003'; $stor4 = '3002'; $stor5 = '3003'; $stor6 = '3004'; $stor7 = '3005'; $stor8 = '3006'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000'; 

	if($type_test eq 1){	
		&query_summary;	
		$counter = 4; 
		&calc8; 

		$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
		$worksheet->merge_range( $a-$counter, 3, $a, 3, $div_name, $border2 );	
	} 
	elsif($type_test eq 3){	
		&query_summary_merchandise;
		$counter = 4; 
		&calc8; 

		$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
		$worksheet->merge_range( $a-$counter, 3, $a, 3, $div_name, $border2 );	
	} 

	$a+=7; 
	$counter = 0; 
	$d=$a;

}


sub new_sheet{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(100);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
#$worksheet->set_print_scale( 100 );
$worksheet->fit_to_pages( 1, 1 );
$worksheet->set_margins( 0.001 );
#$worksheet->conditional_formatting( 'F9:AK6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );
$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 23 );

$worksheet->set_column( 7, 8, 8 );
$worksheet->set_column( 9, 9, 7 );
$worksheet->set_column( 10, 13, undef, undef, 1 );

$worksheet->set_column( 14, 15, 8 );
$worksheet->set_column( 16, 16, 7 );
$worksheet->set_column( 17, 20, undef, undef, 1 );

$worksheet->set_column( 21, 22, 8 );
$worksheet->set_column( 23, 23, 7 );
$worksheet->set_column( 24, 27, undef, undef, 1 );

}

sub new_sheet_2{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(100);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
$worksheet->set_print_scale( 100 );
#$worksheet->fit_to_pages( 1, 1 );
$worksheet->set_margins( 0.001 );
#$worksheet->conditional_formatting( 'F9:AK6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );
$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 23 );

$worksheet->set_column( 7, 8, 8 );
$worksheet->set_column( 9, 9, 7 );
$worksheet->set_column( 10, 13, undef, undef, 1 );

$worksheet->set_column( 14, 15, 8 );
$worksheet->set_column( 16, 16, 7 );
$worksheet->set_column( 17, 20, undef, undef, 1 );

$worksheet->set_column( 21, 22, 8 );
$worksheet->set_column( 23, 23, 7 );
$worksheet->set_column( 24, 27, undef, undef, 1 );

}


sub heading {

$worksheet->write($a-3, 3, "in 000's", $script);
$worksheet->merge_range( $a-3, 7, $a-3, 13, 'OUTRIGHT', $subhead );
$worksheet->merge_range( $a-3, 14, $a-3, 20, 'CONCESSION', $subhead );
$worksheet->merge_range( $a-3, 21, $a-3, 27, 'TOTAL', $subhead );

foreach my $i ( 7, 14, 21 ) {
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

$worksheet->merge_range( $a-1, 7, $a-1, 11, 'General Merchandise', $subhead );
$worksheet->merge_range( $a-1, 12, $a-1, 16, 'Supermarket', $subhead );
$worksheet->merge_range( $a-1, 17, $a-1, 21, 'General Merchandise', $subhead );
$worksheet->merge_range( $a-1, 22, $a-1, 26, 'Supermarket', $subhead );
$worksheet->merge_range( $a-1, 27, $a-1, 31, 'General Merchandise', $subhead );
$worksheet->merge_range( $a-1, 32, $a-1, 36, 'Supermarket', $subhead );

foreach my $i ( 7, 12, 17, 22, 27, 32 ) {
	$worksheet->write($a, $i, "TY", $subhead);
	$worksheet->write($a, $i+1, "LY", $subhead);
	$worksheet->write($a, $i+2, "Growth", $subhead);
	$worksheet->write($a, $i+3, "Budget", $subhead);
	$worksheet->write($a, $i+4, "vs Budget", $subhead);
}

}

# sheet 3
sub query_dept {

$table = 'consolidated_margin.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls1 = $dbh_csv->prepare (qq{SELECT merch_group_code_rev, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
							  FROM $table
							  WHERE new_flg = '$new_flg' or matured_flg = '$matured_flg'
							  GROUP BY merch_group_code_rev
							  ORDER BY merch_group_code_rev
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{merch_group_code_rev};
	#$merch_group_desc = $s->{merch_group_desc};
	
	$sls2 = $dbh_csv->prepare (qq{SELECT group_code, group_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
								 FROM $table 
								 WHERE merch_group_code_rev = '$merch_group_code' and (new_flg = '$new_flg' or matured_flg = '$matured_flg')
								 GROUP BY group_code, group_desc
								 ORDER BY group_code
							 });	
	$sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$group_code = $s->{group_code};
		$group_desc = $s->{group_desc};
				
		$sls3 = $dbh_csv->prepare (qq{SELECT division, division_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
									 FROM $table 
									 WHERE group_code = '$group_code' and merch_group_code_rev = '$merch_group_code' and (new_flg = '$new_flg' or matured_flg = '$matured_flg')
									 GROUP BY division, division_desc
									 ORDER BY division
									});
		$sls3->execute();
		
		$grp_counter = $a;
		while(my $s = $sls3->fetchrow_hashref()){
			$division = $s->{division};
			$division_desc = $s->{division_desc};
			
			$sls4 = $dbh_csv->prepare (qq{SELECT department_code, department_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
										 FROM $table 
										 WHERE group_code = '$group_code' and merch_group_code_rev = '$merch_group_code' and division = '$division' and (new_flg = '$new_flg' or matured_flg = '$matured_flg')
										 GROUP BY department_code, department_desc 
										 ORDER BY department_code
										 });
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){
				
				$worksheet->write($a,5, $s->{department_code},$desc);
				$worksheet->write($a,6, $s->{department_desc},$desc);
				
				$worksheet->write($a,7, $s->{o_retail},$border1);
				$worksheet->write($a,8, $s->{o_margin},$border1);
					if ($s->{o_retail} <= 0){
						$worksheet->write($a,9, "",$subt); }
					else{
						$worksheet->write($a,9, $s->{o_margin}/$s->{o_retail},$subt); }
				
				$worksheet->write($a,10, "",$border1);
				$worksheet->write($a,11, "",$subt);
				$worksheet->write($a,12, "",$border1);
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{c_retail},$border1);
				$worksheet->write($a,15, $s->{c_margin},$border1);
					if ($s->{c_retail} <= 0){
						$worksheet->write($a,16, "",$subt); }
					else{
						$worksheet->write($a,16, $s->{c_margin}/$s->{c_retail},$subt); }
				
				$worksheet->write($a,17, "",$border1);
				$worksheet->write($a,18, "",$subt);
				$worksheet->write($a,19, "",$border1);
				$worksheet->write($a,20, "",$subt);
				
				$worksheet->write($a,21, $s->{o_retail}+$s->{c_retail},$border1);
				$worksheet->write($a,22, $s->{o_margin}+$s->{c_margin},$border1);
					if (($s->{o_retail}+$s->{c_retail}) <= 0){
						$worksheet->write($a,23, "",$subt); }
					else{
						$worksheet->write($a,23, ($s->{o_margin}+$s->{c_margin})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
				$worksheet->write($a,24, "",$border1);
				$worksheet->write($a,25, "",$subt);
				$worksheet->write($a,26, "",$border1);
				$worksheet->write($a,27, "",$subt);					
				
				$a++;
				$counter++;
		
			}
			
			&calc8; #division subtotal
			$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
			$counter = 0; #RESET dept_counter	
			
			$a++; #INCREMENT VARIABLE a
		}

		if($group_code ne 'JW'){
			$grp_o_retail += $s->{o_retail};
			$grp_o_margin += $s->{o_margin};
			
			$grp_c_retail += $s->{c_retail};
			$grp_c_margin += $s->{c_margin};
		}
		
		$worksheet->write($a,7, $s->{o_retail},$border1);
		$worksheet->write($a,8, $s->{o_margin},$border1);
			if ($s->{o_retail} <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, $s->{o_margin}/$s->{o_retail},$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $s->{c_retail},$border1);
		$worksheet->write($a,15, $s->{c_margin},$border1);
			if ($s->{c_retail} <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, $s->{c_margin}/$s->{c_retail},$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);
		
		$worksheet->write($a,21, $s->{o_retail}+$s->{c_retail},$border1);
		$worksheet->write($a,22, $s->{o_margin}+$s->{c_margin},$border1);
			if (($s->{o_retail}+$s->{c_retail}) <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, ($s->{o_margin}+$s->{c_margin})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);				

		$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

		$a++; #INCREMENT VARIABLE a
	}
	
	$total_o_retail += $s->{o_retail};
	$total_o_margin += $s->{o_margin};
	$total_c_retail += $s->{c_retail};
	$total_c_margin += $s->{c_margin};
	
	if ($merch_group_code eq 'DS'){
		############DEPT STORE WO JEWELRY ###############
		$worksheet->write($a,7, $grp_o_retail,$border1);
		$worksheet->write($a,8, $grp_o_margin,$border1);
			if ($grp_o_retail <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, $grp_o_margin/$grp_o_retail,$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $grp_c_retail,$border1);
		$worksheet->write($a,15, $grp_c_margin,$border1);
			if ($grp_c_retail <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, $grp_c_margin/$grp_c_retail,$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);
				
		$worksheet->write($a,21, $grp_o_retail+$grp_c_retail,$border1);
		$worksheet->write($a,22, $grp_o_margin+$grp_c_margin,$border1);
			if (($grp_o_retail+$grp_c_retail) <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, ($grp_o_margin+$grp_c_margin)/($grp_o_retail+$grp_c_retail),$subt); }
				
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);				
		
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store w/o Jewelry', $headD );
	
		$a += 1;
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
	}
	
	elsif($merch_group_code eq 'SU'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
	}
	
	$worksheet->write($a,7, $s->{o_retail},$border1);
	$worksheet->write($a,8, $s->{o_margin},$border1);
		if ($s->{o_retail} <= 0){
			$worksheet->write($a,9, "",$subt); }
		else{
			$worksheet->write($a,9, $s->{o_margin}/$s->{o_retail},$subt); }
				
	$worksheet->write($a,10, "",$border1);
	$worksheet->write($a,11, "",$subt);
	$worksheet->write($a,12, "",$border1);
	$worksheet->write($a,13, "",$subt);
	
	$worksheet->write($a,14, $s->{c_retail},$border1);
	$worksheet->write($a,15, $s->{c_margin},$border1);
		if ($s->{c_retail} <= 0){
			$worksheet->write($a,16, "",$subt); }
		else{
			$worksheet->write($a,16, $s->{c_margin}/$s->{c_retail},$subt); }
				
	$worksheet->write($a,17, "",$border1);
	$worksheet->write($a,18, "",$subt);
	$worksheet->write($a,19, "",$border1);
	$worksheet->write($a,20, "",$subt);
				
	$worksheet->write($a,21, $s->{o_retail}+$s->{c_retail},$border1);
	$worksheet->write($a,22, $s->{o_margin}+$s->{c_margin},$border1);
		if (($s->{o_retail}+$s->{c_retail}) <= 0){
			$worksheet->write($a,23, "",$subt); }
		else{
			$worksheet->write($a,23, ($s->{o_margin}+$s->{c_margin})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
	$worksheet->write($a,24, "",$border1);
	$worksheet->write($a,25, "",$subt);
	$worksheet->write($a,26, "",$border1);
	$worksheet->write($a,27, "",$subt);				
	
	$a++; #INCREMENT VARIABLE a
}

	$worksheet->write($a,7, $total_o_retail,$border1);
	$worksheet->write($a,8, $total_o_margin,$border1);
		if ($total_o_retail <= 0){
			$worksheet->write($a,9, "",$subt); }
		else{
			$worksheet->write($a,9, $total_o_margin/$total_o_retail,$subt); }
	
	$worksheet->write($a,10, "",$border1);
	$worksheet->write($a,11, "",$subt);
	$worksheet->write($a,12, "",$border1);
	$worksheet->write($a,13, "",$subt);
	
	$worksheet->write($a,14, $total_c_retail,$border1);
	$worksheet->write($a,15, $total_c_margin,$border1);
		if ($total_c_retail <= 0){
			$worksheet->write($a,16, "",$subt); }
		else{
			$worksheet->write($a,16, $total_c_margin/$total_c_retail,$subt); }
	
	$worksheet->write($a,17, "",$border1);
	$worksheet->write($a,18, "",$subt);
	$worksheet->write($a,19, "",$border1);
	$worksheet->write($a,20, "",$subt);
	
	$worksheet->write($a,21, $total_o_retail+$total_c_retail,$border1);
	$worksheet->write($a,22, $total_o_margin+$total_c_margin,$border1);
		if (($total_o_retail+$total_c_retail) <= 0){
			$worksheet->write($a,23, "",$subt); }
		else{
			$worksheet->write($a,23, ($total_o_margin+$total_c_margin)/($total_o_retail+$total_c_retail),$subt); }
	
	$worksheet->write($a,24, "",$border1);
	$worksheet->write($a,25, "",$subt);
	$worksheet->write($a,26, "",$border1);
	$worksheet->write($a,27, "",$subt);				
	
$worksheet->write($loc, 2, $loc_desc, $bold);
$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}

sub query_dept_store {

$table = 'consolidated_margin.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls1 = $dbh_csv->prepare (qq{SELECT store_code, store_description, merch_group_code_rev, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
							  FROM $table
							  WHERE store_code = '$store'
							  GROUP BY store_code, store_description, merch_group_code_rev
							  ORDER BY store_code, store_description, merch_group_code_rev
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{merch_group_code_rev};
	$loc_code = $s->{store_code};
	$loc_desc = $s->{store_description};
	
	$sls2 = $dbh_csv->prepare (qq{SELECT group_code, group_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
								 FROM $table 
								 WHERE merch_group_code_rev = '$merch_group_code' and store_code = '$store'
								 GROUP BY group_code, group_desc
								 ORDER BY group_code
							 });	
	$sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$group_code = $s->{group_code};
		$group_desc = $s->{group_desc};
				
		$sls3 = $dbh_csv->prepare (qq{SELECT division, division_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
									 FROM $table 
									 WHERE group_code = '$group_code' and merch_group_code_rev = '$merch_group_code' and store_code = '$store'
									 GROUP BY division, division_desc
									 ORDER BY division
									});
		$sls3->execute();
		
		$grp_counter = $a;
		while(my $s = $sls3->fetchrow_hashref()){
			$division = $s->{division};
			$division_desc = $s->{division_desc};
			
			$sls4 = $dbh_csv->prepare (qq{SELECT department_code, department_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
										 FROM $table 
										 WHERE group_code = '$group_code' and merch_group_code_rev = '$merch_group_code' and division = '$division' and store_code = '$store'
										 GROUP BY department_code, department_desc 
										 ORDER BY department_code
										 });
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){
				
				$worksheet->write($a,5, $s->{department_code},$desc);
				$worksheet->write($a,6, $s->{department_desc},$desc);
				
				$worksheet->write($a,7, $s->{o_retail},$border1);
				$worksheet->write($a,8, $s->{o_margin},$border1);
					if ($s->{o_retail} <= 0){
						$worksheet->write($a,9, "",$subt); }
					else{
						$worksheet->write($a,9, $s->{o_margin}/$s->{o_retail},$subt); }
				
				$worksheet->write($a,10, "",$border1);
				$worksheet->write($a,11, "",$subt);
				$worksheet->write($a,12, "",$border1);
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{c_retail},$border1);
				$worksheet->write($a,15, $s->{c_margin},$border1);
					if ($s->{c_retail} <= 0){
						$worksheet->write($a,16, "",$subt); }
					else{
						$worksheet->write($a,16, $s->{c_margin}/$s->{c_retail},$subt); }
				
				$worksheet->write($a,17, "",$border1);
				$worksheet->write($a,18, "",$subt);
				$worksheet->write($a,19, "",$border1);
				$worksheet->write($a,20, "",$subt);
				
				$worksheet->write($a,21, $s->{o_retail}+$s->{c_retail},$border1);
				$worksheet->write($a,22, $s->{o_margin}+$s->{c_margin},$border1);
					if (($s->{o_retail}+$s->{c_retail}) <= 0){
						$worksheet->write($a,23, "",$subt); }
					else{
						$worksheet->write($a,23, ($s->{o_margin}+$s->{c_margin})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
				$worksheet->write($a,24, "",$border1);
				$worksheet->write($a,25, "",$subt);
				$worksheet->write($a,26, "",$border1);
				$worksheet->write($a,27, "",$subt);					
				
				$a++;
				$counter++;
		
			}
			
			&calc8; #division subtotal
			$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
			$counter = 0; #RESET dept_counter	
			
			$a++; #INCREMENT VARIABLE a
		}

		if($group_code ne 'JW'){
			$grp_o_retail += $s->{o_retail};
			$grp_o_margin += $s->{o_margin};
			
			$grp_c_retail += $s->{c_retail};
			$grp_c_margin += $s->{c_margin};
		}
		
		$worksheet->write($a,7, $s->{o_retail},$border1);
		$worksheet->write($a,8, $s->{o_margin},$border1);
			if ($s->{o_retail} <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, $s->{o_margin}/$s->{o_retail},$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $s->{c_retail},$border1);
		$worksheet->write($a,15, $s->{c_margin},$border1);
			if ($s->{c_retail} <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, $s->{c_margin}/$s->{c_retail},$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);
		
		$worksheet->write($a,21, $s->{o_retail}+$s->{c_retail},$border1);
		$worksheet->write($a,22, $s->{o_margin}+$s->{c_margin},$border1);
			if (($s->{o_retail}+$s->{c_retail}) <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, ($s->{o_margin}+$s->{c_margin})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);				

		$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

		$a++; #INCREMENT VARIABLE a
	}
	
	$total_o_retail += $s->{o_retail};
	$total_o_margin += $s->{o_margin};
	$total_c_retail += $s->{c_retail};
	$total_c_margin += $s->{c_margin};
	
	if ($merch_group_code eq 'DS'){
		############DEPT STORE WO JEWELRY ###############
		$worksheet->write($a,7, $grp_o_retail,$border1);
		$worksheet->write($a,8, $grp_o_margin,$border1);
			if ($grp_o_retail <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, $grp_o_margin/$grp_o_retail,$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $grp_c_retail,$border1);
		$worksheet->write($a,15, $grp_c_margin,$border1);
			if ($grp_c_retail <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, $grp_c_margin/$grp_c_retail,$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);
				
		$worksheet->write($a,21, $grp_o_retail+$grp_c_retail,$border1);
		$worksheet->write($a,22, $grp_o_margin+$grp_c_margin,$border1);
			if (($grp_o_retail+$grp_c_retail) <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, ($grp_o_margin+$grp_c_margin)/($grp_o_retail+$grp_c_retail),$subt); }
				
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);				
		
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store w/o Jewelry', $headD );
	
		$a += 1;
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
	}
	
	elsif($merch_group_code eq 'SU'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
	}
	
	$worksheet->write($a,7, $s->{o_retail},$border1);
	$worksheet->write($a,8, $s->{o_margin},$border1);
		if ($s->{o_retail} <= 0){
			$worksheet->write($a,9, "",$subt); }
		else{
			$worksheet->write($a,9, $s->{o_margin}/$s->{o_retail},$subt); }
				
	$worksheet->write($a,10, "",$border1);
	$worksheet->write($a,11, "",$subt);
	$worksheet->write($a,12, "",$border1);
	$worksheet->write($a,13, "",$subt);
	
	$worksheet->write($a,14, $s->{c_retail},$border1);
	$worksheet->write($a,15, $s->{c_margin},$border1);
		if ($s->{c_retail} <= 0){
			$worksheet->write($a,16, "",$subt); }
		else{
			$worksheet->write($a,16, $s->{c_margin}/$s->{c_retail},$subt); }
				
	$worksheet->write($a,17, "",$border1);
	$worksheet->write($a,18, "",$subt);
	$worksheet->write($a,19, "",$border1);
	$worksheet->write($a,20, "",$subt);
				
	$worksheet->write($a,21, $s->{o_retail}+$s->{c_retail},$border1);
	$worksheet->write($a,22, $s->{o_margin}+$s->{c_margin},$border1);
		if (($s->{o_retail}+$s->{c_retail}) <= 0){
			$worksheet->write($a,23, "",$subt); }
		else{
			$worksheet->write($a,23, ($s->{o_margin}+$s->{c_margin})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
	$worksheet->write($a,24, "",$border1);
	$worksheet->write($a,25, "",$subt);
	$worksheet->write($a,26, "",$border1);
	$worksheet->write($a,27, "",$subt);				
	
	$a++; #INCREMENT VARIABLE a
}

	$worksheet->write($a,7, $total_o_retail,$border1);
	$worksheet->write($a,8, $total_o_margin,$border1);
		if ($total_o_retail <= 0){
			$worksheet->write($a,9, "",$subt); }
		else{
			$worksheet->write($a,9, $total_o_margin/$total_o_retail,$subt); }
	
	$worksheet->write($a,10, "",$border1);
	$worksheet->write($a,11, "",$subt);
	$worksheet->write($a,12, "",$border1);
	$worksheet->write($a,13, "",$subt);
	
	$worksheet->write($a,14, $total_c_retail,$border1);
	$worksheet->write($a,15, $total_c_margin,$border1);
		if ($total_c_retail <= 0){
			$worksheet->write($a,16, "",$subt); }
		else{
			$worksheet->write($a,16, $total_c_margin/$total_c_retail,$subt); }
	
	$worksheet->write($a,17, "",$border1);
	$worksheet->write($a,18, "",$subt);
	$worksheet->write($a,19, "",$border1);
	$worksheet->write($a,20, "",$subt);
	
	$worksheet->write($a,21, $total_o_retail+$total_c_retail,$border1);
	$worksheet->write($a,22, $total_o_margin+$total_c_margin,$border1);
		if (($total_o_retail+$total_c_retail) <= 0){
			$worksheet->write($a,23, "",$subt); }
		else{
			$worksheet->write($a,23, ($total_o_margin+$total_c_margin)/($total_o_retail+$total_c_retail),$subt); }
	
	$worksheet->write($a,24, "",$border1);
	$worksheet->write($a,25, "",$subt);
	$worksheet->write($a,26, "",$border1);
	$worksheet->write($a,27, "",$subt);	

$worksheet->write($loc, 2, $loc_code . " - " . $loc_desc, $bold);			
$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}

# sheet 1
sub query_summary {

$table = 'consolidated_margin.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls = $dbh_csv->prepare (qq{SELECT SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (department_code = '$dept_grp1' or department_code = '$dept_grp2' 
																										or department_code = '$dept_grp3' or department_code = '$dept_grp4' 
																										or department_code = '$dept_grp5' or department_code = '$dept_grp6' 
																										or department_code = '$dept_grp7')) or
																		division = '$division_grp3'))
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40')
								});
$sls->execute();


while(my $s = $sls->fetchrow_hashref()){
	$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
	$worksheet->write($a,7, $s->{o_retail},$border1);
	$worksheet->write($a,8, $s->{o_margin},$border1);
		if ($s->{o_retail} <= 0){
			$worksheet->write($a,9, "",$subt); }
		else{
			$worksheet->write($a,9, $s->{o_margin}/$s->{o_retail},$subt); }
	
	$worksheet->write($a,10, "",$border1);
	$worksheet->write($a,11, "",$subt);
	$worksheet->write($a,12, "",$border1);
	$worksheet->write($a,13, "",$subt);
	
	$worksheet->write($a,14, $s->{c_retail},$border1);
	$worksheet->write($a,15, $s->{c_margin},$border1);
		if ($s->{c_retail} <= 0){
			$worksheet->write($a,16, "",$subt); }
		else{
			$worksheet->write($a,16, $s->{c_margin}/$s->{c_retail},$subt); }
	
	$worksheet->write($a,17, "",$border1);
	$worksheet->write($a,18, "",$subt);
	$worksheet->write($a,19, "",$border1);
	$worksheet->write($a,20, "",$subt);
		
	$worksheet->write($a,21, $s->{o_retail}+$s->{c_retail},$border1);
	$worksheet->write($a,22, $s->{o_margin}+$s->{c_margin},$border1);
		if (($s->{o_retail}+$s->{c_retail}) <= 0){
			$worksheet->write($a,23, "",$subt); }
		else{
			$worksheet->write($a,23, ($s->{o_margin}+$s->{c_margin})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
	$worksheet->write($a,24, "",$border1);
	$worksheet->write($a,25, "",$subt);
	$worksheet->write($a,26, "",$border1);
	$worksheet->write($a,27, "",$subt);		
				
	$a++;
	$counter++;
}

$sls->finish();

}

sub query_by_store {

$table = 'consolidated_margin.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;


$sls = $dbh_csv->prepare (qq{SELECT store_code, store_description, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (department_code = '$dept_grp1' or department_code = '$dept_grp2' 
																										or department_code = '$dept_grp3' or department_code = '$dept_grp4' 
																										or department_code = '$dept_grp5' or department_code = '$dept_grp6' 
																										or department_code = '$dept_grp7')) or
																		division = '$division_grp3')) 
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40')
								 GROUP BY store_code, store_description 
								 ORDER BY store_code
								});
$sls->execute();

while(my $s = $sls->fetchrow_hashref()){
	
	if($s1f2_counter ne 2){
		$worksheet->write($a,5, $s->{store_code},$desc);
		$worksheet->write($a,6, $s->{store_description},$desc);
		$worksheet->write($a,7, $s->{o_retail},$border1);
		$worksheet->write($a,8, $s->{o_margin},$border1);
			if ($s->{o_retail} <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, $s->{o_margin}/$s->{o_retail},$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $s->{c_retail},$border1);
		$worksheet->write($a,15, $s->{c_margin},$border1);
			if ($s->{c_retail} <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, $s->{c_margin}/$s->{c_retail},$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);
			
		$worksheet->write($a,21, $s->{o_retail}+$s->{c_retail},$border1);
		$worksheet->write($a,22, $s->{o_margin}+$s->{c_margin},$border1);
			if (($s->{o_retail}+$s->{c_retail}) <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, ($s->{o_margin}+$s->{c_margin})/($s->{o_retail}+$s->{c_retail}),$subt); }
					
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);
		
			if ($mrch1 eq 'SU' and $mrch2 eq 'SU' and ($s->{store_code} eq '2001' or $s->{store_code} eq '2001W')) {
				$s1f2_o_retail += $s->{o_retail};
				$s1f2_o_margin += $s->{o_margin};
				$s1f2_c_retail += $s->{c_retail};
				$s1f2_c_margin += $s->{c_margin};
				$s1f2_counter ++; # once value = 2, we'll have a summation of s1 and f2
			}
		
		$a++;
		$counter++;
		
	}
	
	if($s1f2_counter eq 2){
		$worksheet->write($a,5, "",$desc);
		$worksheet->write($a,6, "METRO COLON + F2",$desc);
		$worksheet->write($a,7, $s1f2_o_retail,$border1);
		$worksheet->write($a,8, $s1f2_o_margin,$border1);
			if ($s1f2_o_retail <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, $s1f2_o_margin/$s1f2_o_retail,$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $s1f2_c_retail,$border1);
		$worksheet->write($a,15, $s1f2_c_margin,$border1);
			if ($s1f2_c_retail <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, $s1f2_c_margin/$s1f2_c_retail,$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);
			
		$worksheet->write($a,21, $s1f2_o_retail+$s1f2_c_retail,$border1);
		$worksheet->write($a,22, $s1f2_o_margin+$s1f2_c_margin,$border1);
			if (($s1f2_o_retail+$s1f2_c_retail) <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, ($s1f2_o_margin+$s1f2_c_margin)/($s1f2_o_retail+$s1f2_c_retail),$subt); }
					
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);
		
		$worksheet->set_row( $a, undef, undef, 1, undef, undef ); #we hide this row
		
		$s1f2_row = $a;
		$s1f2_counter = 0;
		$a++;
		$counter++;
	}
	
}

$sls->finish();

}

# sheet 2
sub query_summary_merchandise {

$table = 'consolidated_margin.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls = $dbh_csv->prepare (qq{SELECT store_description, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (department_code = '$dept_grp1' or department_code = '$dept_grp2' 
																										or department_code = '$dept_grp3' or department_code = '$dept_grp4' 
																										or department_code = '$dept_grp5' or department_code = '$dept_grp6' 
																										or department_code = '$dept_grp7')) or
																		division = '$division_grp3'))
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40')
								});
$sls->execute();

if ($mrch1 eq 'DS' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){

	while(my $s = $sls->fetchrow_hashref()){
		$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
		$worksheet->write($a,7, $s->{o_retail},$border1);
		$worksheet->write($a,8, $s->{o_margin},$border1);
			if ($s->{o_retail} <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, $s->{o_margin}/$s->{o_retail},$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $s->{c_retail},$border1);
		$worksheet->write($a,15, $s->{c_margin},$border1);
			if ($s->{c_retail} <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, $s->{c_margin}/$s->{c_retail},$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);
			
		$worksheet->write($a,21, $s->{o_retail}+$s->{c_retail},$border1);
		$worksheet->write($a,22, $s->{o_margin}+$s->{c_margin},$border1);
			if (($s->{o_retail}+$s->{c_retail}) <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, ($s->{o_margin}+$s->{c_margin})/($s->{o_retail}+$s->{c_retail}),$subt); }
					
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);
		
		$a++;
		$counter++;
	}
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'SU' and $mrch3 eq 'OT' ){

	while(my $s = $sls->fetchrow_hashref()){
		$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
		$worksheet->write($a,7, "",$border1);
		$worksheet->write($a,8, "",$border1);
		$worksheet->write($a,9, "",$border1);
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$border1);
		
		$worksheet->write($a,12, $s->{wtd_net_ty},$border1);
		$worksheet->write($a,13, $s->{wtd_net_ly},$border1);
		$worksheet->write($a,15, $s->{wtd_target},$border1);
		
			if ($s->{wtd_net_ly} <= 0){ 		$worksheet->write($a,14, "",$subt); 	}
			else{ 		$worksheet->write($a,14, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); 	}
			
			if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0){ 		$worksheet->write($a,16, "",$subt); 	}
			else{ 		$worksheet->write($a,16, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); 	}
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$border1);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$border1);
		$worksheet->write($a,21, "",$border1);
		
		$worksheet->write($a,22, $s->{mtd_net_ty},$border1);
		$worksheet->write($a,23, $s->{mtd_net_ly},$border1);
		$worksheet->write($a,25, $s->{mtd_target},$border1);
		
			if ($s->{mtd_net_ly} <= 0){ 		$worksheet->write($a,24, "",$subt); 	}
			else{ 		$worksheet->write($a,24, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); 	}
			
			if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0){ 		$worksheet->write($a,26, "",$subt); 	}
			else{ 		$worksheet->write($a,26, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); 	}
			
		$worksheet->write($a,27, "",$border1);
		$worksheet->write($a,28, "",$border1);
		$worksheet->write($a,29, "",$border1);
		$worksheet->write($a,30, "",$border1);
		$worksheet->write($a,31, "",$border1);
		
		$worksheet->write($a,32, $s->{qtd_net_ty},$border1);
		$worksheet->write($a,33, $s->{qtd_net_ly},$border1);
		$worksheet->write($a,35, $s->{qtd_target},$border1);
		
			if ($s->{qtd_net_ly} <= 0){ 		$worksheet->write($a,34, "",$subt); 	}
			else{ 		$worksheet->write($a,34, ($s->{qtd_net_ty}-$s->{qtd_net_ly})/$s->{qtd_net_ly},$subt); 	}
			
			if ($s->{qtd_net_ty} <= 0 or $s->{qtd_target} <= 0){ 		$worksheet->write($a,36, "",$subt); 	}
			else{ 		$worksheet->write($a,36, ($s->{qtd_net_ty}-$s->{qtd_target})/$s->{qtd_target},$subt); 	}
		
		$a++;
		$counter++;
	}
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){
	$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
	$worksheet->write($a,7, "",$border1);
	$worksheet->write($a,8, "",$border1);
	$worksheet->write($a,9, "",$border1);
	$worksheet->write($a,10, "",$border1);
	$worksheet->write($a,11, "",$border1);
	$worksheet->write($a,12, "",$border1);
	$worksheet->write($a,13, "",$border1);
	$worksheet->write($a,14, "",$border1);
	$worksheet->write($a,15, "",$border1);
	$worksheet->write($a,16, "",$border1);
	$worksheet->write($a,17, "",$border1);
	$worksheet->write($a,18, "",$border1);
	$worksheet->write($a,19, "",$border1);
	$worksheet->write($a,20, "",$border1);
	$worksheet->write($a,21, "",$border1);
	$worksheet->write($a,22, "",$border1);
	$worksheet->write($a,23, "",$border1);
	$worksheet->write($a,24, "",$border1);
	$worksheet->write($a,25, "",$border1);
	$worksheet->write($a,26, "",$border1);
	$worksheet->write($a,27, "",$border1);
	$worksheet->write($a,28, "",$border1);
	$worksheet->write($a,29, "",$border1);
	$worksheet->write($a,30, "",$border1);
	$worksheet->write($a,31, "",$border1);
	$worksheet->write($a,32, "",$border1);
	$worksheet->write($a,33, "",$border1);
	$worksheet->write($a,34, "",$border1);
	$worksheet->write($a,35, "",$border1);
	$worksheet->write($a,36, "",$border1);
	
	$sls_2 = $dbh_csv->prepare (qq{SELECT merch_group_code_rev, store_description, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (department_code = '$dept_grp1' or department_code = '$dept_grp2' 
																										or department_code = '$dept_grp3' or department_code = '$dept_grp4' 
																										or department_code = '$dept_grp5' or department_code = '$dept_grp6' 
																										or department_code = '$dept_grp7')) or
																		division = '$division_grp3'))
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40')
								GROUP BY merch_group_code_rev
								ORDER BY merch_group_code_rev
								});
	$sls_2->execute();

	while(my $s = $sls_2->fetchrow_hashref()){
	
		if($s->{merch_group_code_rev} eq 'DS'){
			$worksheet->write($a,7, $s->{wtd_net_ty},$border1);
			$worksheet->write($a,8, $s->{wtd_net_ly},$border1);
			$worksheet->write($a,10, $s->{wtd_target},$border1);
			
				if ($s->{wtd_net_ly} <= 0){ 		$worksheet->write($a,9, "",$subt); 	}
				else{ 		$worksheet->write($a,9, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); 	}
				
				if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0){ 		$worksheet->write($a,11, "",$subt); 	}
				else{ 		$worksheet->write($a,11, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); 	}
			
			$worksheet->write($a,17, $s->{mtd_net_ty},$border1);
			$worksheet->write($a,18, $s->{mtd_net_ly},$border1);
			$worksheet->write($a,20, $s->{mtd_target},$border1);
			
				if ($s->{mtd_net_ly} <= 0){ 		$worksheet->write($a,19, "",$subt); 	}
				else{ 		$worksheet->write($a,19, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); 	}
				
				if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0){ 		$worksheet->write($a,21, "",$subt); 	}
				else{ 		$worksheet->write($a,21, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); 	}
				
			$worksheet->write($a,27, $s->{qtd_net_ty},$border1);
			$worksheet->write($a,28, $s->{qtd_net_ly},$border1);
			$worksheet->write($a,30, $s->{qtd_target},$border1);
			
				if ($s->{qtd_net_ly} <= 0){ 		$worksheet->write($a,29, "",$subt); 	}
				else{ 		$worksheet->write($a,29, ($s->{qtd_net_ty}-$s->{qtd_net_ly})/$s->{qtd_net_ly},$subt); 	}
				
				if ($s->{qtd_net_ty} <= 0 or $s->{qtd_target} <= 0){ 		$worksheet->write($a,31, "",$subt); 	}
				else{ 		$worksheet->write($a,31, ($s->{qtd_net_ty}-$s->{qtd_target})/$s->{qtd_target},$subt); 	}
		}
		
		else{		
			$worksheet->write($a,12, $s->{wtd_net_ty},$border1);
			$worksheet->write($a,13, $s->{wtd_net_ly},$border1);
			$worksheet->write($a,15, $s->{wtd_target},$border1);
			
				if ($s->{wtd_net_ly} <= 0){ 		$worksheet->write($a,14, "",$subt); 	}
				else{ 		$worksheet->write($a,14, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); 	}
				
				if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0){ 		$worksheet->write($a,16, "",$subt); 	}
				else{ 		$worksheet->write($a,16, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); 	}
			
			$worksheet->write($a,22, $s->{mtd_net_ty},$border1);
			$worksheet->write($a,23, $s->{mtd_net_ly},$border1);
			$worksheet->write($a,25, $s->{mtd_target},$border1);
			
				if ($s->{mtd_net_ly} <= 0){ 		$worksheet->write($a,24, "",$subt); 	}
				else{ 		$worksheet->write($a,24, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); 	}
				
				if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0){ 		$worksheet->write($a,26, "",$subt); 	}
				else{ 		$worksheet->write($a,26, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); 	}		
				
			$worksheet->write($a,32, $s->{qtd_net_ty},$border1);
			$worksheet->write($a,33, $s->{qtd_net_ly},$border1);
			$worksheet->write($a,35, $s->{qtd_target},$border1);
			
				if ($s->{qtd_net_ly} <= 0){ 		$worksheet->write($a,34, "",$subt); 	}
				else{ 		$worksheet->write($a,34, ($s->{qtd_net_ty}-$s->{qtd_net_ly})/$s->{qtd_net_ly},$subt); 	}
				
				if ($s->{qtd_net_ty} <= 0 or $s->{qtd_target} <= 0){ 		$worksheet->write($a,36, "",$subt); 	}
				else{ 		$worksheet->write($a,36, ($s->{qtd_net_ty}-$s->{qtd_target})/$s->{qtd_target},$subt); 	}		
		}
		
	}
	
	$sls_2->finish();
	
	$a++;
	$counter++;
}

$sls->finish();

}

sub query_by_store_merchandise {

$table = 'consolidated_margin.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;

$blank = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, hidden => 1 );
$worksheet->conditional_formatting( 'H44:AF60', { type     => 'cell',  criteria => '=', value    => 0, format   => $blank });	
$worksheet->conditional_formatting( 'F9:AK2000',  { type => 'cell', criteria => '<', value => 0, format => $down } );			

$sls = $dbh_csv->prepare (qq{SELECT store_code, store_description, merch_group_code_rev, store_description, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (department_code = '$dept_grp1' or department_code = '$dept_grp2' 
																										or department_code = '$dept_grp3' or department_code = '$dept_grp4' 
																										or department_code = '$dept_grp5' or department_code = '$dept_grp6' 
																										or department_code = '$dept_grp7')) or
																		division = '$division_grp3')) 
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40')
								 GROUP BY store_code, store_description, merch_group_code_rev
								 ORDER BY store_code, merch_group_code_rev
								});
$sls->execute();

if ($mrch1 eq 'DS' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){	
	while(my $s = $sls->fetchrow_hashref()){
			$worksheet->write($a,5, $s->{store_code},$desc);
			$worksheet->write($a,6, $s->{store_description},$desc);
			
			if($s->{merch_group_code_rev} eq 'DS'){
				$worksheet->write($a,7, $s->{wtd_net_ty},$border1);
				$worksheet->write($a,8, $s->{wtd_net_ly},$border1);
				$worksheet->write($a,10, $s->{wtd_target},$border1);
				
					if ($s->{wtd_net_ly} <= 0){ 				$worksheet->write($a,9, "",$subt); }
					else{ 				$worksheet->write($a,9, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); }
					
					if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0 ){ 				$worksheet->write($a,11, "",$subt); }
					else{ 				$worksheet->write($a,11, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); }
				
				$worksheet->write($a,12, "",$border1);
				$worksheet->write($a,13, "",$border1);
				$worksheet->write($a,14, "",$border1);
				$worksheet->write($a,15, "",$border1);
				$worksheet->write($a,16, "",$border1);
					
				$worksheet->write($a,17, $s->{mtd_net_ty},$border1);
				$worksheet->write($a,18, $s->{mtd_net_ly},$border1);
				$worksheet->write($a,20, $s->{mtd_target},$border1);
				
					if ($s->{mtd_net_ly} <= 0){ 				$worksheet->write($a,19, "",$subt); }
					else{ 				$worksheet->write($a,19, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); }
					
					if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0 ){ 				$worksheet->write($a,21, "",$subt); }
					else{ 				$worksheet->write($a,21, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); }
				
				$worksheet->write($a,22, "",$border1);
				$worksheet->write($a,23, "",$border1);
				$worksheet->write($a,24, "",$border1);
				$worksheet->write($a,25, "",$border1);
				$worksheet->write($a,26, "",$border1);
				
				$worksheet->write($a,27, $s->{qtd_net_ty},$border1);
				$worksheet->write($a,28, $s->{qtd_net_ly},$border1);
				$worksheet->write($a,30, $s->{qtd_target},$border1);
				
					if ($s->{qtd_net_ly} <= 0){ 				$worksheet->write($a,29, "",$subt); }
					else{ 				$worksheet->write($a,29, ($s->{qtd_net_ty}-$s->{qtd_net_ly})/$s->{qtd_net_ly},$subt); }
					
					if ($s->{qtd_net_ty} <= 0 or $s->{qtd_target} <= 0 ){ 				$worksheet->write($a,31, "",$subt); }
					else{ 				$worksheet->write($a,31, ($s->{qtd_net_ty}-$s->{qtd_target})/$s->{qtd_target},$subt); }
				
				$worksheet->write($a,32, "",$border1);
				$worksheet->write($a,33, "",$border1);
				$worksheet->write($a,34, "",$border1);
				$worksheet->write($a,35, "",$border1);
				$worksheet->write($a,36, "",$border1);
			}
			
			else{	
				$a -= 1;
				$counter -= 1;
							
				$worksheet->write($a,12, $s->{wtd_net_ty},$border1);
				$worksheet->write($a,13, $s->{wtd_net_ly},$border1);
				$worksheet->write($a,15, $s->{wtd_target},$border1);
				
					if ($s->{wtd_net_ly} <= 0){ 				$worksheet->write($a,14, "",$subt); }
					else{ 				$worksheet->write($a,14, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); }
					
					if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0 ){ 				$worksheet->write($a,16, "",$subt); }
					else{ 				$worksheet->write($a,16, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); }
									
				$worksheet->write($a,22, $s->{mtd_net_ty},$border1);
				$worksheet->write($a,23, $s->{mtd_net_ly},$border1);
				$worksheet->write($a,25, $s->{mtd_target},$border1);
				
					if ($s->{mtd_net_ly} <= 0){ 				$worksheet->write($a,24, "",$subt); }
					else{ 				$worksheet->write($a,24, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); }
					
					if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0 ){ 				$worksheet->write($a,21, "",$subt); }
					else{ 				$worksheet->write($a,26, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); }
					
				$worksheet->write($a,32, $s->{qtd_net_ty},$border1);
				$worksheet->write($a,33, $s->{qtd_net_ly},$border1);
				$worksheet->write($a,35, $s->{qtd_target},$border1);
				
					if ($s->{qtd_net_ly} <= 0){ 				$worksheet->write($a,34, "",$subt); }
					else{ 				$worksheet->write($a,34, ($s->{qtd_net_ty}-$s->{qtd_net_ly})/$s->{qtd_net_ly},$subt); }
					
					if ($s->{qtd_net_ty} <= 0 or $s->{qtd_target} <= 0 ){ 				$worksheet->write($a,36, "",$subt); }
					else{ 				$worksheet->write($a,36, ($s->{qtd_net_ty}-$s->{qtd_target})/$s->{qtd_target},$subt); }
			}
			
			$a++;
			$counter++;			
	}
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'SU' and $mrch3 eq 'OT' ){	
	
	while(my $s = $sls->fetchrow_hashref()){
		
		if($s1f2_counter ne 2){
			$worksheet->write($a,5, $s->{store_code},$desc);
			$worksheet->write($a,6, $s->{store_description},$desc);
			
			if($s->{merch_group_code_rev} eq 'DS'){ 		
				$worksheet->write($a,7, $s->{wtd_net_ty},$border1);
				$worksheet->write($a,8, $s->{wtd_net_ly},$border1);
				$worksheet->write($a,10, $s->{wtd_target},$border1);
				
					if ($s->{wtd_net_ly} <= 0){ 				$worksheet->write($a,9, "",$subt); }
					else{ 				$worksheet->write($a,9, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); }
					
					if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0 ){ 				$worksheet->write($a,11, "",$subt); }
					else{ 				$worksheet->write($a,11, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); }
									
				$worksheet->write($a,17, $s->{mtd_net_ty},$border1);
				$worksheet->write($a,18, $s->{mtd_net_ly},$border1);
				$worksheet->write($a,20, $s->{mtd_target},$border1);
				
					if ($s->{mtd_net_ly} <= 0){ 				$worksheet->write($a,19, "",$subt); }
					else{ 				$worksheet->write($a,19, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); }
					
					if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0 ){ 				$worksheet->write($a,21, "",$subt); }
					else{ 				$worksheet->write($a,21, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); }	

				$worksheet->write($a,27, $s->{qtd_net_ty},$border1);
				$worksheet->write($a,28, $s->{qtd_net_ly},$border1);
				$worksheet->write($a,30, $s->{qtd_target},$border1);
				
					if ($s->{qtd_net_ly} <= 0){ 				$worksheet->write($a,29, "",$subt); }
					else{ 				$worksheet->write($a,29, ($s->{qtd_net_ty}-$s->{qtd_net_ly})/$s->{qtd_net_ly},$subt); }
					
					if ($s->{qtd_net_ty} <= 0 or $s->{qtd_target} <= 0 ){ 				$worksheet->write($a,31, "",$subt); }
					else{ 				$worksheet->write($a,31, ($s->{qtd_net_ty}-$s->{qtd_target})/$s->{qtd_target},$subt); }	

				$a -= 1;
				$counter -= 1;	
			}
			
			else{						
				$worksheet->write($a,12, $s->{wtd_net_ty},$border1);
				$worksheet->write($a,13, $s->{wtd_net_ly},$border1);
				$worksheet->write($a,15, $s->{wtd_target},$border1);
				
					if ($s->{wtd_net_ly} <= 0){ 				$worksheet->write($a,14, "",$subt); }
					else{ 				$worksheet->write($a,14, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); }
					
					if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0 ){ 				$worksheet->write($a,16, "",$subt); }
					else{ 				$worksheet->write($a,16, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); }
					
				$worksheet->write($a,22, $s->{mtd_net_ty},$border1);
				$worksheet->write($a,23, $s->{mtd_net_ly},$border1);
				$worksheet->write($a,25, $s->{mtd_target},$border1);
				
					if ($s->{mtd_net_ly} <= 0){ 				$worksheet->write($a,24, "",$subt); }
					else{ 				$worksheet->write($a,24, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); }
					
					if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0 ){ 				$worksheet->write($a,26, "",$subt); }
					else{ 				$worksheet->write($a,26, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); }
					
				$worksheet->write($a,32, $s->{qtd_net_ty},$border1);
				$worksheet->write($a,33, $s->{qtd_net_ly},$border1);
				$worksheet->write($a,35, $s->{qtd_target},$border1);
				
					if ($s->{qtd_net_ly} <= 0){ 				$worksheet->write($a,34, "",$subt); }
					else{ 				$worksheet->write($a,34, ($s->{qtd_net_ty}-$s->{qtd_net_ly})/$s->{qtd_net_ly},$subt); }
					
					if ($s->{qtd_net_ty} <= 0 or $s->{qtd_target} <= 0 ){ 				$worksheet->write($a,36, "",$subt); }
					else{ 				$worksheet->write($a,36, ($s->{qtd_net_ty}-$s->{qtd_target})/$s->{qtd_target},$subt); }
				
					if ($s->{merch_group_code_rev} eq 'SU' and ($s->{store_code} eq '2001' or $s->{store_code} eq '2001W')) {
						$s1f2_wtd_net_ty += $s->{wtd_net_ty};
						$s1f2_wtd_net_ly += $s->{wtd_net_ly};
						$s1f2_wtd_target += $s->{wtd_target};
						$s1f2_mtd_net_ty += $s->{mtd_net_ty};
						$s1f2_mtd_net_ly += $s->{mtd_net_ly};
						$s1f2_mtd_target += $s->{mtd_target};
						$s1f2_qtd_net_ty += $s->{qtd_net_ty};
						$s1f2_qtd_net_ly += $s->{qtd_net_ly};
						$s1f2_qtd_target += $s->{qtd_target};
						$s1f2_counter ++; # once value = 2, we'll have a summation of s1 and f2						
					}
			}		
			
			$a++;
			$counter++;			
		}
		
		if($s1f2_counter eq 2){
			$worksheet->write($a,5, "",$desc);
			$worksheet->write($a,6, "METRO COLON + F2",$desc);
			
			$worksheet->write($a,12, $s1f2_wtd_net_ty,$border1);
			$worksheet->write($a,13, $s1f2_wtd_net_ly,$border1);
			$worksheet->write($a,15, $s1f2_wtd_target,$border1);
			
			if ($s1f2_wtd_net_ly <= 0){ 				$worksheet->write($a,14, "",$subt); }
			else{ 				$worksheet->write($a,14, ($s1f2_wtd_net_ty-$s1f2_wtd_net_ly)/$s1f2_wtd_net_ly,$subt); }
			
			if ($s1f2_wtd_net_ty <= 0 or $s1f2_wtd_target <= 0 ){ 				$worksheet->write($a,16, "",$subt); }
			else{ 				$worksheet->write($a,16, ($s1f2_wtd_net_ty-$s1f2_wtd_target)/$s1f2_wtd_target,$subt); }
				
			$worksheet->write($a,22, $s1f2_mtd_net_ty,$border1);
			$worksheet->write($a,23, $s1f2_mtd_net_ly,$border1);
			$worksheet->write($a,25, $s1f2_mtd_target,$border1);
			
			if ($s1f2_mtd_net_ly <= 0){ 				$worksheet->write($a,24, "",$subt); }
			else{ 				$worksheet->write($a,24, ($s1f2_mtd_net_ty-$s1f2_mtd_net_ly)/$s1f2_mtd_net_ly,$subt); }
			
			if ($s1f2_mtd_net_ty <= 0 or $s1f2_mtd_target <= 0 ){ 				$worksheet->write($a,26, "",$subt); }
			else{ 				$worksheet->write($a,26, ($s1f2_mtd_net_ty-$s1f2_mtd_target)/$s1f2_mtd_target,$subt); }
			
			$worksheet->write($a,32, $s1f2_qtd_net_ty,$border1);
			$worksheet->write($a,33, $s1f2_qtd_net_ly,$border1);
			$worksheet->write($a,35, $s1f2_qtd_target,$border1);
			
			if ($s1f2_qtd_net_ly <= 0){ 				$worksheet->write($a,34, "",$subt); }
			else{ 				$worksheet->write($a,34, ($s1f2_qtd_net_ty-$s1f2_qtd_net_ly)/$s1f2_qtd_net_ly,$subt); }
			
			if ($s1f2_qtd_net_ty <= 0 or $s1f2_qtd_target <= 0 ){ 				$worksheet->write($a,36, "",$subt); }
			else{ 				$worksheet->write($a,36, ($s1f2_qtd_net_ty-$s1f2_qtd_target)/$s1f2_qtd_target,$subt); }
			
			$worksheet->set_row( $a, undef, undef, 1, undef, undef ); #we hide this row
			
			$s1f2_row = $a;
			$s1f2_counter = 0;
			
			$a++;
			$counter++;
		}	
		}
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){
	
	$sls_2 = $dbh_csv->prepare (qq{SELECT store_code, store_description, merch_group_code_rev, store_description, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (department_code = '$dept_grp1' or department_code = '$dept_grp2' 
																										or department_code = '$dept_grp3' or department_code = '$dept_grp4' 
																										or department_code = '$dept_grp5' or department_code = '$dept_grp6' 
																										or department_code = '$dept_grp7')) or
																		division = '$division_grp3')) 
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40')
								 GROUP BY store_code, store_description , merch_group_code_rev
								 ORDER BY store_code, merch_group_code_rev
								});
	$sls_2->execute();	

	while(my $s = $sls_2->fetchrow_hashref()){
	
	$worksheet->write($a,5, $s->{store_code},$desc);
	$worksheet->write($a,6, $s->{store_description},$desc);
	
		if($s->{merch_group_code_rev} eq 'DS'){
			$worksheet->write($a,7, $s->{wtd_net_ty},$border1);
			$worksheet->write($a,8, $s->{wtd_net_ly},$border1);
			$worksheet->write($a,10, $s->{wtd_target},$border1);
			
				if ($s->{wtd_net_ly} <= 0){ 		$worksheet->write($a,9, "",$subt); 	}
				else{ 		$worksheet->write($a,9, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); 	}
				
				if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0){ 		$worksheet->write($a,11, "",$subt); 	}
				else{ 		$worksheet->write($a,11, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); 	}
			
			$worksheet->write($a,17, $s->{mtd_net_ty},$border1);
			$worksheet->write($a,18, $s->{mtd_net_ly},$border1);
			$worksheet->write($a,20, $s->{mtd_target},$border1);
			
				if ($s->{mtd_net_ly} <= 0){ 		$worksheet->write($a,19, "",$subt); 	}
				else{ 		$worksheet->write($a,19, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); 	}
				
				if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0){ 		$worksheet->write($a,21, "",$subt); 	}
				else{ 		$worksheet->write($a,21, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); 	}
				
			$worksheet->write($a,27, $s->{qtd_net_ty},$border1);
			$worksheet->write($a,28, $s->{qtd_net_ly},$border1);
			$worksheet->write($a,30, $s->{qtd_target},$border1);
			
				if ($s->{qtd_net_ly} <= 0){ 		$worksheet->write($a,29, "",$subt); 	}
				else{ 		$worksheet->write($a,29, ($s->{qtd_net_ty}-$s->{qtd_net_ly})/$s->{qtd_net_ly},$subt); 	}
				
				if ($s->{qtd_net_ty} <= 0 or $s->{qtd_target} <= 0){ 		$worksheet->write($a,31, "",$subt); 	}
				else{ 		$worksheet->write($a,31, ($s->{qtd_net_ty}-$s->{qtd_target})/$s->{qtd_target},$subt); 	}
		}
		
		else{	
			$worksheet->write($a,12, $s->{wtd_net_ty},$border1);
			$worksheet->write($a,13, $s->{wtd_net_ly},$border1);
			$worksheet->write($a,15, $s->{wtd_target},$border1);
			
				if ($s->{wtd_net_ly} <= 0){ 		$worksheet->write($a,14, "",$subt); 	}
				else{ 		$worksheet->write($a,14, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); 	}
				
				if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0){ 		$worksheet->write($a,16, "",$subt); 	}
				else{ 		$worksheet->write($a,16, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); 	}
			
			$worksheet->write($a,22, $s->{mtd_net_ty},$border1);
			$worksheet->write($a,23, $s->{mtd_net_ly},$border1);
			$worksheet->write($a,25, $s->{mtd_target},$border1);
			
				if ($s->{mtd_net_ly} <= 0){ 		$worksheet->write($a,24, "",$subt); 	}
				else{ 		$worksheet->write($a,24, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); 	}
				
				if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0){ 		$worksheet->write($a,26, "",$subt); 	}
				else{ 		$worksheet->write($a,26, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); 	}

			$worksheet->write($a,32, $s->{qtd_net_ty},$border1);
			$worksheet->write($a,33, $s->{qtd_net_ly},$border1);
			$worksheet->write($a,35, $s->{qtd_target},$border1);
			
				if ($s->{qtd_net_ly} <= 0){ 		$worksheet->write($a,34, "",$subt); 	}
				else{ 		$worksheet->write($a,34, ($s->{qtd_net_ty}-$s->{qtd_net_ly})/$s->{qtd_net_ly},$subt); 	}
				
				if ($s->{qtd_net_ty} <= 0 or $s->{qtd_target} <= 0){ 		$worksheet->write($a,36, "",$subt); 	}
				else{ 		$worksheet->write($a,36, ($s->{qtd_net_ty}-$s->{qtd_target})/$s->{qtd_target},$subt); 	}	
		
			$a++;
			$counter++;
		}
		
	}
	
	$sls_2->finish();
}

$sls->finish();

}


sub calc8 { 

if($type_test eq 3 or $type_test eq 4){
	foreach my $col( 7, 8, 10, 12, 13, 15, 17, 18, 20, 22, 23, 25, 27, 28, 30, 32, 33, 35 ){
		if($s1f2_row ne 0){
			my $sum = '=(SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . '))-'. xl_rowcol_to_cell( $s1f2_row, $col );
				$worksheet->write( $a, $col, $sum, $bodyNum );
			
			if($col eq 8 or $col eq 13 or $col eq 18 or $col eq 23 or $col eq 28 or $col eq 33){
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );	
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
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );	
			}
			elsif ($col eq 10 or $col eq 15 or $col eq 20 or $col eq 25 or $col eq 30 or $col eq 35){
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
			}
		}		
	}
}

else{
	foreach my $col( 7, 8, 10, 12, 14, 15, 17, 19, 21, 22, 24, 26 ){
		if($s1f2_row ne 0){
			my $sum = '=(SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . '))-'. xl_rowcol_to_cell( $s1f2_row, $col );
			$worksheet->write( $a, $col, $sum, $bodyNum );}
		
		else{
			my $sum = '=SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . ')';
				$worksheet->write( $a, $col, $sum, $bodyNum );	}
			
			if ($col eq 8 or $col eq 15 or $col eq 22){
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
			elsif ($col eq 10 or $col eq 17 or $col eq 24){
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
			elsif ($col eq 12 or $col eq 19 or $col eq 26){
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} } } }


sub generate_csv {

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
# my $uid = 'kent';
# my $pw = 'amer1c8';
my $uid = 'rmsprd';
my $pw = 'noida123';

# my $hostname = "10.128.4.23";
# my $sid = "MGRMST";
# my $port = '1521';
# my $uid = 'rmsprd';
# my $pw = 'vicsal123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh, ">", "consolidated_margin.csv" or die "consolidated_margin.csv: $!";

$test = qq{ 
SELECT 
SGD.STORE_FORMAT,
CASE WHEN TO_CHAR(SGD.STORE) = '4002' THEN '2001W' ELSE TO_CHAR(SGD.STORE) END AS STORE_CODE,
CASE WHEN SGD.STORE IN ('2012', '2013', '3009', '4004', '3010', '3011') THEN 'SU' || SGD.STORE	 WHEN SGD.STORE IN ('4002') THEN 'SU2001W'     WHEN SGD.STORE = '2223' THEN 'DS' || SGD.STORE	 ELSE SGD.MERCH_GROUP_CODE || SGD.STORE END AS STORE,
SGD.STORE_NAME STORE_DESCRIPTION, 
CASE WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 9000) THEN 'DS'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8500) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO = 8040) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO != 8040) THEN 'DS'ELSE SGD.MERCH_GROUP_CODE END AS MERCH_GROUP_CODE_REV,
SGD.MERCH_GROUP_CODE, 
SGD.MERCH_GROUP_DESC, 
SGD.ATTRIB1 GROUP_CODE, 
SGD.ATTRIB2 GROUP_DESC, 
SGD.DIVISION, 
SGD.DIV_NAME DIVISION_DESC, 
SGD.GROUP_NO DEPARTMENT_CODE, 
SGD.GROUP_NAME DEPARTMENT_DESC, 
(SUM(OTR.OTR_COST)/1000) OTR_COST,
(SUM(OTR.OTR_RETAIL)/1000) OTR_RETAIL,
(SUM(OTR.OTR_MARGIN)/1000) OTR_MARGIN, 
(SUM(CON.CON_COST)/1000) CON_COST,
(SUM(CON.CON_RETAIL)/1000) CON_RETAIL,
(SUM(CON.CON_MARGIN)/1000) CON_MARGIN,
CASE WHEN SGD.STORE IN ('2009','2012','7176','7003','2006','7004','2007','6008','7005','2010','7300','7009','5006','5005','5004','5003','5002','5001','7173','3004','3003','3001','4003','7008','7007','7006','3006','3007','2003','7000','4002','3005','3002','2002','2001','2011','2008','7001','2004','7002','2005') THEN 0 ELSE 1 END AS NEW_FLG,
CASE WHEN SGD.STORE IN ('2009','2012','7176','7003','2006','7004','2007','6008','7005','2010','7300','7009','5006','5005','5004','5003','5002','5001','7173','3004','3003','3001','4003','7008','7007','7006','3006','3007','2003','7000','4002','3005','3002','2002','2001','2011','2008','7001','2004','7002','2005') THEN 1 ELSE 0 END AS MATURED_FLG
FROM
	(SELECT S.STORE_FORMAT, S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
		FROM
			(SELECT DISTINCT STORE_FORMAT, STORE, STORE_NAME FROM STORE UNION ALL SELECT DISTINCT 8 AS STORE_FORMAT, WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
			(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
				FROM DEPS D 
				  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
				  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
				  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
	GROUP BY S.STORE_FORMAT, S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
	)SGD
	
	LEFT JOIN
	(SELECT T.LOCATION, T.GROUP_NO, SUM(T.UNITS) OTR_UNITS, SUM(T.OTR_COST) OTR_COST, SUM(T.OTR_RETAIL) OTR_RETAIL, SUM((NVL(T.OTR_RETAIL,0))-(NVL(T.TOTAL_COST,0))) OTR_MARGIN FROM
		(SELECT TDH.LOCATION, DEPS.GROUP_NO, TDH.UNITS UNITS, TDH.TOTAL_COST OTR_COST, TDH.TOTAL_RETAIL OTR_RETAIL, 
			TDH.TOTAL_COST AS TOTAL_COST
			--CASE WHEN TDH.NEW_TOTAL_COST IS NULL THEN TDH.TOTAL_COST ELSE TDH.NEW_TOTAL_COST END AS TOTAL_COST
			--FROM TRAN_DATA_HISTORY_08042014 TDH 
			FROM TRAN_DATA_HISTORY TDH 
			--FROM TRAN_DATA_HISTORY_BUP TDH 
				JOIN DEPS ON TDH.DEPT = DEPS.DEPT
			WHERE (TDH.TRAN_DATE BETWEEN '01-OCT-13' AND '31-OCT-13') AND TDH.TRAN_CODE = 1 AND DEPS.PURCHASE_TYPE = 0)T
	GROUP BY T.LOCATION, T.GROUP_NO)OTR
	ON SGD.STORE = OTR.LOCATION AND SGD.GROUP_NO = OTR.GROUP_NO
	
	LEFT JOIN
	(SELECT T.LOCATION, T.GROUP_NO, SUM(T.UNITS) CON_UNITS, SUM(T.CON_COST) CON_COST, SUM(T.CON_RETAIL) CON_RETAIL, SUM((NVL(T.CON_RETAIL,0))-(NVL(T.TOTAL_COST,0))) CON_MARGIN FROM
		(SELECT TDH.LOCATION, DEPS.GROUP_NO, TDH.UNITS UNITS, TDH.TOTAL_COST CON_COST, TDH.TOTAL_RETAIL CON_RETAIL, 
			TDH.TOTAL_COST AS TOTAL_COST
			--CASE WHEN TDH.NEW_TOTAL_COST IS NULL THEN TDH.TOTAL_COST ELSE TDH.NEW_TOTAL_COST END AS TOTAL_COST
			--FROM TRAN_DATA_HISTORY_08042014 TDH 
			FROM TRAN_DATA_HISTORY TDH 
			--FROM TRAN_DATA_HISTORY_BUP TDH 
				JOIN DEPS ON TDH.DEPT = DEPS.DEPT
			WHERE (TDH.TRAN_DATE BETWEEN '01-OCT-13' AND '31-OCT-13') AND TDH.TRAN_CODE = 1 AND DEPS.PURCHASE_TYPE = 2)T
	GROUP BY T.LOCATION, T.GROUP_NO)CON
	ON SGD.STORE = CON.LOCATION AND SGD.GROUP_NO = CON.GROUP_NO
	
GROUP BY 
SGD.STORE_FORMAT,
CASE WHEN TO_CHAR(SGD.STORE) = '4002' THEN '2001W' ELSE TO_CHAR(SGD.STORE) END,
CASE WHEN SGD.STORE IN ('2012', '2013', '3009', '4004', '3010', '3011') THEN 'SU' || SGD.STORE 	 WHEN SGD.STORE IN ('4002') THEN 'SU2001W'     WHEN SGD.STORE = '2223' THEN 'DS' || SGD.STORE	 ELSE SGD.MERCH_GROUP_CODE || SGD.STORE END,
SGD.STORE_NAME, 
CASE WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 9000) THEN 'DS'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8500) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO = 8040) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO != 8040) THEN 'DS'ELSE SGD.MERCH_GROUP_CODE END,
SGD.MERCH_GROUP_CODE, 
SGD.MERCH_GROUP_DESC, 
SGD.ATTRIB1, 
SGD.ATTRIB2, 
SGD.DIVISION, 
SGD.DIV_NAME, 
SGD.GROUP_NO, 
SGD.GROUP_NAME,
CASE WHEN SGD.STORE IN ('2009','2012','7176','7003','2006','7004','2007','6008','7005','2010','7300','7009','5006','5005','5004','5003','5002','5001','7173','3004','3003','3001','4003','7008','7007','7006','3006','3007','2003','7000','4002','3005','3002','2002','2001','2011','2008','7001','2004','7002','2005') THEN 0 ELSE 1 END,
CASE WHEN SGD.STORE IN ('2009','2012','7176','7003','2006','7004','2007','6008','7005','2010','7300','7009','5006','5005','5004','5003','5002','5001','7173','3004','3003','3001','4003','7008','7007','7006','3006','3007','2003','7000','4002','3005','3002','2002','2001','2011','2008','7001','2004','7002','2005') THEN 1 ELSE 0 END
ORDER BY 1, 3, 5, 7, 9};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_lc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "consolidated_margin.csv: $!";
 
$sth->finish();
$dbh->disconnect;

}

# mailer
sub mail_grp1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1, $attachment_file_2 ) = @ARGV;

$to = ' gerry.guanlao@metrogaisano.com, eric.redona@metrogaisano.com, lucille.malazarte@metrogaisano.com, tricia.luntao@metrogaisano.com, jj.moreno@metrogaisano.com, cj.jesena@metrogaisano.com, rex.cabanilla@metrogaisano.com ';

$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, christopher.calalang@metrogaisano.com, annalyn.conde@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com ';
		
# $to = ' annalyn.conde@metrogaisano.com ';
# $bcc = 'kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, christopher.calalang@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com';
		
$subject = 'Daily Sales Performance as of ' . $as_of;
$msgbody_file = 'message_BI.txt';

$attachment_file_1 = "Daily Sales Performance - Summary (as of $as_of) v1.3.xlsx";
$attachment_file_2 = "Daily Sales Performance - Summary (as of $as_of) v1.3.pdf";

my $msgbody = read_file( $msgbody_file );

my $attachment_data_1 = encode_base64( read_file( $attachment_file_1, 1 ));
my $attachment_data_2 = encode_base64( read_file( $attachment_file_2, 1 ));

my %mail = (
    To   => $to,
    Subject => $subject
);

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

sub mail_grp2 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1, $attachment_file_2 ) = @ARGV;

# $to = ' gerry.guanlao@metrogaisano.com, ernest.potgieter@metrogaisano.com, emily.silverio@metrogaisano.com, cj.jesena@metrogaisano.com, joefrey.camu@metrogaisano.com, may.sasedor@metrogaisano.com, alain.reyes@metrogaisano.com, dinah.ramirez@metrogaisano.com, jacqueline.cano@metrogaisano.com, limuel.ulanday@metrogaisano.com, glenda.navares@metrogaisano.com, roy.igot@metrogaisano.com  ';
# $bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, christopher.calalang@metrogaisano.com, annalyn.conde@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com ';
		
$to = 'fnaquines@metro.com.ph';
$bcc = 'kent.mamalias@metrogaisano.com';
		
$subject = 'Daily Sales Performance as of ' . $as_of;
$msgbody_file = 'message_BI.txt';

$attachment_file_1 = "Daily Sales Performance (as of $as_of).xlsx";
$attachment_file_2 = "Daily Sales Performance (as of $as_of).pdf";

my $msgbody = read_file( $msgbody_file );

my $attachment_data_1 = encode_base64( read_file( $attachment_file_1, 1 ));
my $attachment_data_2 = encode_base64( read_file( $attachment_file_2, 1 ));

my %mail = (
    To   => $to,
    Subject => $subject
);

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















