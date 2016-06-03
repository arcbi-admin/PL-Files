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
 
	$workbook = Excel::Writer::XLSX->new("CONSOLIDATED MARGIN - Summary (as of $as_of) v1.3.xlsx");
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

	&new_sheet($sheet = "GenMerch_Spmkt");
	&call_str_merchandise;
	
	&new_sheet_2($sheet = "Department");			
	&call_div;
		
	$workbook->close();
	
	&mail_grp1;	
	
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

$grp_o_retail = 0, $grp_o_margin = 0, $grp_c_retail = 0, $grp_c_margin = 0, $grp_amt432000 = 0, $grp_amt433000 = 0, $grp_amt458490 = 0, $grp_amt434000 = 0, $grp_amt458550 = 0, $grp_amt460100 = 0, $grp_amt460200 = 0, $grp_amt460300 = 0, $grp_amt503200 = 0, $grp_amt503200 = 0, $grp_amt503500 = 0, $grp_amt506000 = 0, $grp_amt501000 = 0;

$total_o_retail = 0, $total_o_margin = 0, $total_c_retail = 0, $total_c_margin = 0, $total_amt432000 = 0, $total_amt433000 = 0, $total_amt458490 = 0, $total_amt434000 = 0, $total_amt458550 = 0, $total_amt460100 = 0, $total_amt460200 = 0, $total_amt460300  = 0, $total_amt503200 = 0, $total_amt503250 = 0, $total_amt503500 = 0, $total_amt506000 = 0, $total_amt501000 = 0;

$type_test = 0;

$worksheet->write($a-9, 3, "Total Margin Performance", $bold1);
#$worksheet->write($a-8, 3, "MTD: $mo_st_date_fld - $mo_en_date_fld");
$worksheet->write($a-8, 3, "MTD: 30 Nov 2013 - 30 Nov 2013");
$worksheet->write($a-7, 3, "As of $as_of");

##========================= COMP STORES ===========================##

&heading_2;
&heading;
&query_dept($new_flg = 0, $matured_flg = 1, $loc_desc = "COMP STORES");

##========================= ALL STORES ===========================##

$a += 7;

$grp_o_retail = 0, $grp_o_margin = 0, $grp_c_retail = 0, $grp_c_margin = 0, $grp_amt432000 = 0, $grp_amt433000 = 0, $grp_amt458490 = 0, $grp_amt434000 = 0, $grp_amt458550 = 0, $grp_amt460100 = 0, $grp_amt460200 = 0, $grp_amt460300 = 0, $grp_amt503200 = 0, $grp_amt503200 = 0, $grp_amt503500 = 0, $grp_amt506000 = 0, $grp_amt501000 = 0;

$total_o_retail = 0, $total_o_margin = 0, $total_c_retail = 0, $total_c_margin = 0, $total_amt432000 = 0, $total_amt433000 = 0, $total_amt458490 = 0, $total_amt434000 = 0, $total_amt458550 = 0, $total_amt460100 = 0, $total_amt460200 = 0, $total_amt460300  = 0, $total_amt503200 = 0, $total_amt503250 = 0, $total_amt503500 = 0, $total_amt506000 = 0, $total_amt501000 = 0;

$type_test = 0;

&heading_2;
&heading;
&query_dept($new_flg = 1, $matured_flg = 1, $loc_desc = "ALL STORES");

##========================= BY STORE ===========================##

foreach my $i ( '2001', '2001W', '2002', '2003', '2004', '2005', '2006', '2007', '2008', '2009', '2010', '2011', '2012', '2013', '2223', '3001', '3002', '3003', '3004', '3005', '3006', '3007', '3008', '3009', '3010', '3011', '3012', '3013', '4003', '4004', '6001', '6002', '6003', '6004', '6005', '6006', '6009', '6010', '6011', '6012', '6013' ){ 
# foreach my $i ( '2001', '2001W', '2002' ){ 

	$a += 7;
	$grp_o_retail = 0, $grp_o_margin = 0, $grp_c_retail = 0, $grp_c_margin = 0, $grp_amt432000 = 0, $grp_amt433000 = 0, $grp_amt458490 = 0, $grp_amt434000 = 0, $grp_amt458550 = 0, $grp_amt460100 = 0, $grp_amt460200 = 0, $grp_amt460300 = 0, $grp_amt503200 = 0, $grp_amt503200 = 0, $grp_amt503500 = 0, $grp_amt506000 = 0, $grp_amt501000 = 0;

	$total_o_retail = 0, $total_o_margin = 0, $total_c_retail = 0, $total_c_margin = 0, $total_amt432000 = 0, $total_amt433000 = 0, $total_amt458490 = 0, $total_amt434000 = 0, $total_amt458550 = 0, $total_amt460100 = 0, $total_amt460200 = 0, $total_amt460300  = 0, $total_amt503200 = 0, $total_amt503250 = 0, $total_amt503500 = 0, $total_amt506000 = 0, $total_amt501000 = 0;

	&heading_2;
	&heading;
	&query_dept_store($store = $i);

}

}

sub call_str {

$a=10, $d=0, $e=10, $counter=0, $comp_su = 0, $new_su = 0, $comp_nb = 0, $comp_hy = 0, $new_hy = 0, $comp_ds = 0, $new_ds = 0, $type_test = 1;

$worksheet->write($a-10, 3, "Total Margin Performance", $bold1);
$worksheet->write($a-9, 3, "MTD: 30 Nov 2013 - 30 Nov 2013");
$worksheet->write($a-8, 3, "As of $as_of");

$worksheet->write($a-6, 3, "Summary", $bold);

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

$worksheet->write($a-6, 3, "Per Store", $bold);

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
&strNew_Nb;

}

sub call_str_merchandise {

$a=10, $d=0, $e=10, $counter=0, $comp_su = 0, $new_su = 0, $comp_nb = 0, $comp_hy = 0, $new_hy = 0, $comp_ds = 0, $new_ds = 0, $type_test = 3;

$worksheet->write($a-9, 3, "Total Margin Performance", $bold1);
$worksheet->write($a-8, 3, "MTD: 30 Nov 2013 - 30 Nov 2013");
$worksheet->write($a-7, 3, "As of $as_of");

$worksheet->write($a-4, 3, "Summary", $bold);

&heading_3;

$worksheet->merge_range( $a-3, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-3, 4, $a-1, 6, 'Format', $subhead );

#$a += 1;

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

$worksheet->merge_range( $a-3, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-3, 4, $a-1, 4, 'Type', $subhead );
$worksheet->merge_range( $a-3, 5, $a-1, 5, 'Code', $subhead );
$worksheet->merge_range( $a-3, 6, $a-1, 6, 'Desc', $subhead );

#$a += 1;

&strComp_Ds;
&strNew_Ds;
&strComp_Su;
&strNew_Su;
&strComp_Hy;
&strNew_Hy;
&strComp_Nb;
&strNew_Nb;

}


sub strComp_Su {

$div_name = "Comp";  $div_name3 = "Supermarket";
$mrch1 = 'SU'; $mrch2 = 'SU'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8040'; $dept_grp2 = '0000'; $dept_grp3 = '0000'; $dept_grp4 = '0000'; $dept_grp5 = '0000'; $dept_grp6 = '0000'; $dept_grp7 = '0000';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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

			foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_su,$col).','.xl_rowcol_to_cell($new_su,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94){
						my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $var, $bodyPct );	} }				

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

			foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94,96,97,99,101,103,104,106,108,110,111,113,115,117,118,120,121,123,124,126,127,129,130,132,133,135,136,138,139,141,142,144,145,147,148,150,151,153,154,156,157,159,161,163,164,166,168,170,171,173,174,176,177,179,180,182,183 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_su,$col).','.xl_rowcol_to_cell($new_su,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75 or $col eq 97 or $col eq 104 or $col eq 111 or $col eq 157 or $col eq 164){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77 or $col eq 99 or $col eq 106 or $col eq 113 or $col eq 159 or $col eq 168){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79 or $col eq 101 or $col eq 108 or $col eq 115 or $col eq 161 or $col eq 168){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94 
					  or $col eq 118 or $col eq 121 or $col eq 124 or $col eq 127 or $col eq 130 or $col eq 133 or $col eq 136 or $col eq 139 or $col eq 142 or $col eq 145 or $col eq 148 or $col eq 151 or $col eq 154 or $col eq 171 or $col eq 174 or $col eq 177 or $col eq 180 or $col eq 183){
						my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $var, $bodyPct );	}
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
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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

		$tst = $a-$counter; $comp_nb=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94){
						my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $var, $bodyPct );	} }	

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

		$tst = $a-$counter; $comp_nb=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER
	
	}
	
}

sub strNew_Nb {

$div_name = "New"; $div_name2 = "Neighborhood";  $div_name3 = "Neighborhood Store";
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '0000'; $division_grp2 = '0000';  $division_grp3 = '0000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
$s1f2_counter = 0, $s1f2_row = 0;

$store1 = 'SU3012'; $store2 = 'DS3012'; $store3 = 'OT3012'; $store4 = '0000'; $store5 = '0000'; $store6 = '0000'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '3012'; $stor2 = '0000'; $stor3 = '0000'; $stor4 = '0000'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000';  

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
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		#$tst = $a; 
		$new_nb=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).','.xl_rowcol_to_cell($new_nb,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94){
						my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $var, $bodyPct );	} }	

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
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_nb=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94,96,97,99,101,103,104,106,108,110,111,113,115,117,118,120,121,123,124,126,127,129,130,132,133,135,136,138,139,141,142,144,145,147,148,150,151,153,154,156,157,159,161,163,164,166,168,170,171,173,174,176,177,179,180,182,183 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).','.xl_rowcol_to_cell($new_nb,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75 or $col eq 97 or $col eq 104 or $col eq 111 or $col eq 157 or $col eq 164){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77 or $col eq 99 or $col eq 106 or $col eq 113 or $col eq 159 or $col eq 168){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79 or $col eq 101 or $col eq 108 or $col eq 115 or $col eq 161 or $col eq 168){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94 
					  or $col eq 118 or $col eq 121 or $col eq 124 or $col eq 127 or $col eq 130 or $col eq 133 or $col eq 136 or $col eq 139 or $col eq 142 or $col eq 145 or $col eq 148 or $col eq 151 or $col eq 154 or $col eq 171 or $col eq 174 or $col eq 177 or $col eq 180 or $col eq 183){
						my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $var, $bodyPct );	}
			}

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name2, $bodyN );
		$a+=1; $counter = 0; $d=$a;
		
	}
	
}

sub strComp_Hy {

$div_name = "Comp";  $div_name3 = "Hypermarket";
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_o_retail = 0, $s1f2_o_margin = 0, $s1f2_c_retail = 0, $s1f2_c_margin = 0;
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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

			foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_hy,$col).','.xl_rowcol_to_cell($new_hy,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94){
						my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $var, $bodyPct );	} }	

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

			foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94,96,97,99,101,103,104,106,108,110,111,113,115,117,118,120,121,123,124,126,127,129,130,132,133,135,136,138,139,141,142,144,145,147,148,150,151,153,154,156,157,159,161,163,164,166,168,170,171,173,174,176,177,179,180,182,183 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_hy,$col).','.xl_rowcol_to_cell($new_hy,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75 or $col eq 97 or $col eq 104 or $col eq 111 or $col eq 157 or $col eq 164){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77 or $col eq 99 or $col eq 106 or $col eq 113 or $col eq 159 or $col eq 168){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79 or $col eq 101 or $col eq 108 or $col eq 115 or $col eq 161 or $col eq 168){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94 
					  or $col eq 118 or $col eq 121 or $col eq 124 or $col eq 127 or $col eq 130 or $col eq 133 or $col eq 136 or $col eq 139 or $col eq 142 or $col eq 145 or $col eq 148 or $col eq 151 or $col eq 154 or $col eq 171 or $col eq 174 or $col eq 177 or $col eq 180 or $col eq 183){
						my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $var, $bodyPct );	}
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
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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

			foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_ds,$col).','.xl_rowcol_to_cell($new_ds,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94){
						my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $var, $bodyPct );	} }	

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

			foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94,96,97,99,101,103,104,106,108,110,111,113,115,117,118,120,121,123,124,126,127,129,130,132,133,135,136,138,139,141,142,144,145,147,148,150,151,153,154,156,157,159,161,163,164,166,168,170,171,173,174,176,177,179,180,182,183 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_ds,$col).','.xl_rowcol_to_cell($new_ds,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75 or $col eq 97 or $col eq 104 or $col eq 111 or $col eq 157 or $col eq 164){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77 or $col eq 99 or $col eq 106 or $col eq 113 or $col eq 159 or $col eq 168){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79 or $col eq 101 or $col eq 108 or $col eq 115 or $col eq 161 or $col eq 168){
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
					elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94 
					  or $col eq 118 or $col eq 121 or $col eq 124 or $col eq 127 or $col eq 130 or $col eq 133 or $col eq 136 or $col eq 139 or $col eq 142 or $col eq 145 or $col eq 148 or $col eq 151 or $col eq 154 or $col eq 171 or $col eq 174 or $col eq 177 or $col eq 180 or $col eq 183){
						my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $var, $bodyPct );	}
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
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
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
$s1f2_amt432000 = 0, $s1f2_amt433000 = 0, $s1f2_amt458490 = 0, $s1f2_amt434000 = 0, $s1f2_amt458550 = 0, $s1f2_amt460100 = 0, $s1f2_amt460200 = 0, $s1f2_amt460300 = 0, $s1f2_amt503200 = 0, $s1f2_amt503250 = 0, $s1f2_amt503500 = 0, $s1f2_amt506000 = 0, $s1f2_amt501000 = 0;
$s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU3001'; $store2 = 'SU3007'; $store3 = 'SU4003'; $store4 = 'DS3001'; $store5 = 'DS3007'; $store6 = 'DS4003'; $store7 = 'OT3001'; $store8 = 'OT3007'; $store9 = 'OT4003'; $store10 = 'SU3002'; $store11 = 'DS3002'; $store12 = 'OT3002'; $store13 = 'SU3003'; $store14 = 'DS3003'; $store15 = 'OT3003'; $store16 = 'SU3004'; $store17 = 'DS3004'; $store18 = 'OT3004'; $store19 = 'SU3005'; $store20 = 'DS3005'; $store21 = 'OT3005'; $store22 = 'SU3006'; $store23 = 'DS3006'; $store24 = 'OT3006'; $store25 = 'SU3012'; $store26 = 'DS3012';  $store27 = 'OT3012'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '30001'; $stor2 = '3007'; $stor3 = '4003'; $stor4 = '3002'; $stor5 = '3003'; $stor6 = '3004'; $stor7 = '3005'; $stor8 = '3006'; $stor9 = '3012'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000'; 

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
$worksheet->set_zoom(90);
$worksheet->set_paper(8);
$worksheet->center_horizontally();
$worksheet->fit_to_pages( 1, 1 );
$worksheet->set_margins( 0.001 );
#$worksheet->conditional_formatting( 'F9:AK6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );

}

sub new_sheet_2{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(90);
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

$worksheet->write($a-5, 3, "in 000's", $script);
$worksheet->merge_range( $a-4, 7, $a-3, 13, 'TOTAL', $subhead );
$worksheet->merge_range( $a-4, 14, $a-4, 66, 'OUTRIGHT', $subhead );
$worksheet->merge_range( $a-4, 67, $a-4, 95, 'CONCESSION', $subhead );

$worksheet->merge_range( $a-3, 14, $a-3, 20, 'TOTAL - OUTRIGHT', $subhead );
$worksheet->merge_range( $a-3, 21, $a-3, 27, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 28, $a-3, 54, 'FRONT - OTHERS', $subhead );
$worksheet->merge_range( $a-3, 55, $a-3, 63, 'BACK', $subhead );
$worksheet->merge_range( $a-3, 64, $a-3, 66, 'OTHER COST', $subhead );
$worksheet->merge_range( $a-3, 67, $a-3, 73, 'TOTAL - CONCESSION', $subhead );
$worksheet->merge_range( $a-3, 74, $a-3, 80, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 81, $a-3, 95, 'BACK', $subhead );

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

foreach my $i ( 28, 31, 34, 37, 40, 43, 46, 49, 52, 55, 58, 61, 64, 81, 84, 87, 90, 93 ) {
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
$worksheet->merge_range( $a-2, 81, $a-2, 83, 'Ad Support', $subhead );
$worksheet->merge_range( $a-2, 84, $a-2, 86, 'Other Income-Storage/Concession', $subhead );
$worksheet->merge_range( $a-2, 87, $a-2, 89, 'Light Recovery', $subhead );
$worksheet->merge_range( $a-2, 90, $a-2, 92, 'Water Recovery', $subhead );
$worksheet->merge_range( $a-2, 93, $a-2, 95, 'Supplies Recovery', $subhead );

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

$worksheet->write($a-5, 3, "in 000's", $script);

$worksheet->merge_range( $a-5, 7, $a-5, 95, 'GENERAL MERCHANDISE', $subhead );
$worksheet->merge_range( $a-5, 96, $a-5, 184, 'SUPERMARKET', $subhead );

$worksheet->merge_range( $a-4, 7, $a-3, 13, 'TOTAL', $subhead );
$worksheet->merge_range( $a-4, 14, $a-4, 66, 'OUTRIGHT', $subhead );
$worksheet->merge_range( $a-4, 67, $a-4, 95, 'CONCESSION', $subhead );

$worksheet->merge_range( $a-4, 96, $a-3, 102, 'TOTAL', $subhead );
$worksheet->merge_range( $a-4, 103, $a-4, 155, 'OUTRIGHT', $subhead );
$worksheet->merge_range( $a-4, 156, $a-4, 184, 'CONCESSION', $subhead );

$worksheet->merge_range( $a-3, 14, $a-3, 20, 'TOTAL - OUTRIGHT', $subhead );
$worksheet->merge_range( $a-3, 21, $a-3, 27, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 28, $a-3, 54, 'FRONT - OTHERS', $subhead );
$worksheet->merge_range( $a-3, 55, $a-3, 63, 'BACK', $subhead );
$worksheet->merge_range( $a-3, 64, $a-3, 66, 'OTHER COST', $subhead );
$worksheet->merge_range( $a-3, 67, $a-3, 73, 'TOTAL - CONCESSION', $subhead );
$worksheet->merge_range( $a-3, 74, $a-3, 80, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 81, $a-3, 95, 'BACK', $subhead );

$worksheet->merge_range( $a-3, 103, $a-3, 109, 'TOTAL - OUTRIGHT', $subhead );
$worksheet->merge_range( $a-3, 110, $a-3, 116, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 117, $a-3, 143, 'FRONT - OTHERS', $subhead );
$worksheet->merge_range( $a-3, 144, $a-3, 152, 'BACK', $subhead );
$worksheet->merge_range( $a-3, 153, $a-3, 155, 'OTHER COST', $subhead );
$worksheet->merge_range( $a-3, 156, $a-3, 162, 'TOTAL - CONCESSION', $subhead );
$worksheet->merge_range( $a-3, 163, $a-3, 169, 'FRONT', $subhead );
$worksheet->merge_range( $a-3, 170, $a-3, 184, 'BACK', $subhead );

foreach my $i ( 7, 14, 21, 67, 74, 96, 103, 110, 156, 163 ) {
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

foreach my $i ( 28, 31, 34, 37, 40, 43, 46, 49, 52, 55, 58, 61, 64, 81, 84, 87, 90, 93, 117, 120, 123, 126, 129, 132, 135, 138, 141, 144, 147, 150, 153, 170, 173, 176, 179, 182 ) {
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
$worksheet->merge_range( $a-2, 81, $a-2, 83, 'Ad Support', $subhead );
$worksheet->merge_range( $a-2, 84, $a-2, 86, 'Other Income-Storage/Concession', $subhead );
$worksheet->merge_range( $a-2, 87, $a-2, 89, 'Light Recovery', $subhead );
$worksheet->merge_range( $a-2, 90, $a-2, 92, 'Water Recovery', $subhead );
$worksheet->merge_range( $a-2, 93, $a-2, 95, 'Supplies Recovery', $subhead );

$worksheet->merge_range( $a-2, 117, $a-2, 119, 'Cost of Sales - Trade', $subhead );
$worksheet->merge_range( $a-2, 120, $a-2, 122, 'Transfer Discrepancy', $subhead );
$worksheet->merge_range( $a-2, 123, $a-2, 125, 'Promotional Item Charged to Margin', $subhead );
$worksheet->merge_range( $a-2, 126, $a-2, 128, 'Wastage', $subhead );
$worksheet->merge_range( $a-2, 129, $a-2, 131, 'Invoice Price Variance - Trade', $subhead );
$worksheet->merge_range( $a-2, 132, $a-2, 134, 'Shrinkage Cost', $subhead );
$worksheet->merge_range( $a-2, 135, $a-2, 137, 'Cost Variance', $subhead );
$worksheet->merge_range( $a-2, 138, $a-2, 140, 'Synchronization Account', $subhead );
$worksheet->merge_range( $a-2, 141, $a-2, 143, 'Freight Recovery', $subhead );
$worksheet->merge_range( $a-2, 144, $a-2, 146, 'Purchase Allowance', $subhead );
$worksheet->merge_range( $a-2, 147, $a-2, 149, 'Purchase Discouns-Special', $subhead );
$worksheet->merge_range( $a-2, 150, $a-2, 152, 'Other Income-Metro Vendor Portal', $subhead );
$worksheet->merge_range( $a-2, 153, $a-2, 155, 'Freight', $subhead );
$worksheet->merge_range( $a-2, 170, $a-2, 172, 'Ad Support', $subhead );
$worksheet->merge_range( $a-2, 173, $a-2, 175, 'Other Income-Storage/Concession', $subhead );
$worksheet->merge_range( $a-2, 176, $a-2, 178, 'Light Recovery', $subhead );
$worksheet->merge_range( $a-2, 179, $a-2, 181, 'Water Recovery', $subhead );
$worksheet->merge_range( $a-2, 182, $a-2, 184, 'Supplies Recovery', $subhead );

$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 23 );

}

# sheet 3
sub query_dept {

$table = 'consolidated_margin_x.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls1 = $dbh_csv->prepare (qq{SELECT merch_group_code_rev, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300 , SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
							  FROM $table
							  WHERE new_flg = '$new_flg' or matured_flg = '$matured_flg'
							  GROUP BY merch_group_code_rev
							  ORDER BY merch_group_code_rev
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{merch_group_code_rev};
	#$merch_group_desc = $s->{merch_group_desc};
	
	$sls2 = $dbh_csv->prepare (qq{SELECT group_code, group_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
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
				
		$sls3 = $dbh_csv->prepare (qq{SELECT division, division_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
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
			
			$sls4 = $dbh_csv->prepare (qq{SELECT department_code, department_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
										 FROM $table 
										 WHERE group_code = '$group_code' and merch_group_code_rev = '$merch_group_code' and division = '$division' and (new_flg = '$new_flg' or matured_flg = '$matured_flg')
										 GROUP BY department_code, department_desc 
										 ORDER BY department_code
										 });
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){
				
				$worksheet->write($a,5, $s->{department_code},$desc);
				$worksheet->write($a,6, $s->{department_desc},$desc);
				
				$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1);
				$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if (($s->{o_retail}+$s->{c_retail}) le 0){
						$worksheet->write($a,9, "",$subt); }
					else{
						$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
				$worksheet->write($a,10, "",$border1);
				$worksheet->write($a,11, "",$subt);
				$worksheet->write($a,12, "",$border1);
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{o_retail},$border1);
				$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1);
					if ($s->{o_retail} le 0){
						$worksheet->write($a,16, "",$subt); }
					else{
						$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
				
				$worksheet->write($a,17, "",$border1);
				$worksheet->write($a,18, "",$subt);
				$worksheet->write($a,19, "",$border1);
				$worksheet->write($a,20, "",$subt);			
				
				$worksheet->write($a,21, $s->{o_retail},$border1);
				$worksheet->write($a,22, $s->{o_margin},$border1);
					if ($s->{o_retail} le 0){
						$worksheet->write($a,23, "",$subt); }
					else{
						$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
						
				$worksheet->write($a,24, "",$border1);
				$worksheet->write($a,25, "",$subt);
				$worksheet->write($a,26, "",$border1);
				$worksheet->write($a,27, "",$subt);	
				
				$worksheet->write($a,28, $s->{amt501000},$border1);
				$worksheet->write($a,29, "",$border1);
				$worksheet->write($a,30, "",$border1);
				
				$worksheet->write($a,31, $s->{amt503200},$border1);
				$worksheet->write($a,32, "",$border1);
				$worksheet->write($a,33, "",$border1);
				
				$worksheet->write($a,34, $s->{amt503250},$border1);
				$worksheet->write($a,35, "",$border1);
				$worksheet->write($a,36, "",$border1);
				
				$worksheet->write($a,37, $s->{amt503500},$border1);
				$worksheet->write($a,38, "",$border1);
				$worksheet->write($a,39, "",$border1);
				
				$worksheet->write($a,40, $s->{amt506000},$border1);
				$worksheet->write($a,41, "",$border1);
				$worksheet->write($a,42, "",$border1);
				
				$worksheet->write($a,43, $s->{amt503000},$border1);
				$worksheet->write($a,44, "",$border1);
				$worksheet->write($a,45, "",$border1);				
							
				$worksheet->write($a,46, $s->{amt507000},$border1);
				$worksheet->write($a,47, "",$border1);
				$worksheet->write($a,48, "",$border1);
				
				$worksheet->write($a,49, $s->{amt999998},$border1);
				$worksheet->write($a,50, "",$border1);
				$worksheet->write($a,51, "",$border1);
				
				$worksheet->write($a,52, $s->{amt504000},$border1);
				$worksheet->write($a,53, "",$border1);
				$worksheet->write($a,54, "",$border1);
							
				$worksheet->write($a,55, $s->{amt432000},$border1);
				$worksheet->write($a,56, "",$border1);
				$worksheet->write($a,57, "",$border1);
				
				$worksheet->write($a,58, $s->{amt433000},$border1);
				$worksheet->write($a,59, "",$border1);
				$worksheet->write($a,60, "",$border1);
				
				$worksheet->write($a,61, $s->{amt458490},$border1);
				$worksheet->write($a,62, "",$border1);
				$worksheet->write($a,63, "",$border1);
				
				$worksheet->write($a,64, $s->{amt505000},$border1);
				$worksheet->write($a,65, "",$border1);
				$worksheet->write($a,66, "",$border1);
				
				$worksheet->write($a,67, $s->{c_retail},$border1);
				$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,69, "",$subt); }
					else{
						$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
						
				$worksheet->write($a,70, "",$border1);
				$worksheet->write($a,71, "",$subt);
				$worksheet->write($a,72, "",$border1);
				$worksheet->write($a,73, "",$subt);	
				
				$worksheet->write($a,74, $s->{c_retail},$border1);
				$worksheet->write($a,75, $s->{c_margin},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,76, "",$subt); }
					else{
						$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
					
				$worksheet->write($a,77, "",$border1);
				$worksheet->write($a,78, "",$subt);
				$worksheet->write($a,79, "",$border1);
				$worksheet->write($a,80, "",$subt);	
						
				$worksheet->write($a,81, $s->{amt434000},$border1);
				$worksheet->write($a,82, "",$border1);
				$worksheet->write($a,83, "",$border1);
						
				$worksheet->write($a,84, $s->{amt458550},$border1);
				$worksheet->write($a,85, "",$border1);
				$worksheet->write($a,86, "",$border1);
				
				$worksheet->write($a,87, $s->{amt460100},$border1);
				$worksheet->write($a,88, "",$border1);
				$worksheet->write($a,89, "",$border1);
					
				$worksheet->write($a,90, $s->{amt460200},$border1);
				$worksheet->write($a,91, "",$border1);
				$worksheet->write($a,92, "",$border1);
						
				$worksheet->write($a,93, $s->{amt460300},$border1);
				$worksheet->write($a,94, "",$border1);
				$worksheet->write($a,95, "",$border1);			
				
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
			$grp_amt501000 += $s->{amt501000};
			$grp_amt503200 += $s->{amt503200};
			$grp_amt503250 += $s->{amt503250};
			$grp_amt503500 += $s->{amt503500};
			$grp_amt506000 += $s->{amt506000};
			$grp_amt503000 += $s->{amt503000};
			$grp_amt507000 += $s->{amt507000};
			$grp_amt999998 += $s->{amt999998};
			$grp_amt504000 += $s->{amt504000};
			$grp_amt505000 += $s->{amt505000};
			$grp_amt432000 += $s->{amt432000};
			$grp_amt433000 += $s->{amt433000};
			$grp_amt458490 += $s->{amt458490};
			$grp_amt434000 += $s->{amt434000};
			$grp_amt458550 += $s->{amt458550};
			$grp_amt460100 += $s->{amt460100};
			$grp_amt460200 += $s->{amt460200};
			$grp_amt460300 += $s->{amt460300};
		}
		
		$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$bodyNum);
		$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$bodyNum);
			if (($s->{o_retail}+$s->{c_retail}) le 0){
				$worksheet->write($a,9, "",$bodyPct); }
			else{
				$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$bodyPct); }
				
		$worksheet->write($a,10, "",$bodyNum);
		$worksheet->write($a,11, "",$bodyPct);
		$worksheet->write($a,12, "",$bodyNum);
		$worksheet->write($a,13, "",$bodyPct);
				
		$worksheet->write($a,14, $s->{o_retail},$bodyNum);
		$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$bodyNum);
			if ($s->{o_retail} le 0){
				$worksheet->write($a,16, "",$bodyPct); }
			else{
				$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$bodyPct); }
				
		$worksheet->write($a,17, "",$bodyNum);
		$worksheet->write($a,18, "",$bodyPct);
		$worksheet->write($a,19, "",$bodyNum);
		$worksheet->write($a,20, "",$bodyPct);			
				
		$worksheet->write($a,21, $s->{o_retail},$bodyNum);
		$worksheet->write($a,22, $s->{o_margin},$bodyNum);
			if ($s->{o_retail} le 0){
				$worksheet->write($a,23, "",$bodyPct); }
			else{
				$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$bodyPct); }
						
		$worksheet->write($a,24, "",$bodyNum);
		$worksheet->write($a,25, "",$bodyPct);
		$worksheet->write($a,26, "",$bodyNum);
		$worksheet->write($a,27, "",$bodyPct);	
		
		$worksheet->write($a,28, $s->{amt501000},$bodyNum);
		$worksheet->write($a,29, "",$bodyNum);
		$worksheet->write($a,30, "",$bodyNum);
			
		$worksheet->write($a,31, $s->{amt503200},$bodyNum);
		$worksheet->write($a,32, "",$bodyNum);
		$worksheet->write($a,33, "",$bodyNum);
		
		$worksheet->write($a,34, $s->{amt503250},$bodyNum);
		$worksheet->write($a,35, "",$bodyNum);
		$worksheet->write($a,36, "",$bodyNum);
		
		$worksheet->write($a,37, $s->{amt503500},$bodyNum);
		$worksheet->write($a,38, "",$bodyNum);
		$worksheet->write($a,39, "",$bodyNum);
		
		$worksheet->write($a,40, $s->{amt506000},$bodyNum);
		$worksheet->write($a,41, "",$bodyNum);
		$worksheet->write($a,42, "",$bodyNum);
		
		$worksheet->write($a,43, $s->{amt503000},$bodyNum);
		$worksheet->write($a,44, "",$bodyNum);
		$worksheet->write($a,45, "",$bodyNum);
							
		$worksheet->write($a,46, $s->{amt507000},$bodyNum);
		$worksheet->write($a,47, "",$bodyNum);
		$worksheet->write($a,48, "",$bodyNum);
		
		$worksheet->write($a,49, $s->{amt999998},$bodyNum);
		$worksheet->write($a,50, "",$bodyNum);
		$worksheet->write($a,51, "",$bodyNum);
		
		$worksheet->write($a,52, $s->{amt505000},$bodyNum);
		$worksheet->write($a,53, "",$bodyNum);
		$worksheet->write($a,54, "",$bodyNum);
					
		$worksheet->write($a,55, $s->{amt432000},$bodyNum);
		$worksheet->write($a,56, "",$bodyNum);
		$worksheet->write($a,57, "",$bodyNum);
		
		$worksheet->write($a,58, $s->{amt433000},$bodyNum);
		$worksheet->write($a,59, "",$bodyNum);
		$worksheet->write($a,60, "",$bodyNum);
		
		$worksheet->write($a,61, $s->{amt458490},$bodyNum);
		$worksheet->write($a,62, "",$bodyNum);
		$worksheet->write($a,63, "",$bodyNum);
		
		$worksheet->write($a,64, $s->{amt504000},$bodyNum);
		$worksheet->write($a,65, "",$bodyNum);
		$worksheet->write($a,66, "",$bodyNum);
		
		$worksheet->write($a,67, $s->{c_retail},$bodyNum);
		$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$bodyNum);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,69, "",$bodyPct); }
			else{
				$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$bodyPct); }
				
		$worksheet->write($a,70, "",$bodyNum);
		$worksheet->write($a,71, "",$bodyPct);
		$worksheet->write($a,72, "",$bodyNum);
		$worksheet->write($a,73, "",$bodyPct);	
		
		$worksheet->write($a,74, $s->{c_retail},$bodyNum);
		$worksheet->write($a,75, $s->{c_margin},$bodyNum);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,76, "",$bodyPct); }
			else{
				$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$bodyPct); }
			
		$worksheet->write($a,77, "",$bodyNum);
		$worksheet->write($a,78, "",$bodyPct);
		$worksheet->write($a,79, "",$bodyNum);
		$worksheet->write($a,80, "",$bodyPct);	
				
		$worksheet->write($a,81, $s->{amt434000},$bodyNum);
		$worksheet->write($a,82, "",$bodyNum);
		$worksheet->write($a,83, "",$bodyNum);
				
		$worksheet->write($a,84, $s->{amt458550},$bodyNum);
		$worksheet->write($a,85, "",$bodyNum);
		$worksheet->write($a,86, "",$bodyNum);
		
		$worksheet->write($a,87, $s->{amt460100},$bodyNum);
		$worksheet->write($a,88, "",$bodyNum);
		$worksheet->write($a,89, "",$bodyNum);
			
		$worksheet->write($a,90, $s->{amt460200},$bodyNum);
		$worksheet->write($a,91, "",$bodyNum);
		$worksheet->write($a,92, "",$bodyNum);
				
		$worksheet->write($a,93, $s->{amt460300},$bodyNum);
		$worksheet->write($a,94, "",$bodyNum);
		$worksheet->write($a,95, "",$bodyNum);			

		$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

		$a++; #INCREMENT VARIABLE a
	}
	
	$total_o_retail += $s->{o_retail};
	$total_o_margin += $s->{o_margin};
	$total_c_retail += $s->{c_retail};
	$total_c_margin += $s->{c_margin};
	$total_amt501000 += $s->{amt501000};
	$total_amt503200 += $s->{amt503200};
	$total_amt503250 += $s->{amt503250};
	$total_amt503500 += $s->{amt503500};
	$total_amt506000 += $s->{amt506000};
	$total_amt503000 += $s->{amt503000};
	$total_amt507000 += $s->{amt507000};
	$total_amt999998 += $s->{amt999998};
	$total_amt504000 += $s->{amt504000};
	$total_amt505000 += $s->{amt505000};
	$total_amt432000 += $s->{amt432000};
	$total_amt433000 += $s->{amt433000};
	$total_amt458490 += $s->{amt458490};
	$total_amt434000 += $s->{amt434000};
	$total_amt458550 += $s->{amt458550};
	$total_amt460100 += $s->{amt460100};
	$total_amt460200 += $s->{amt460200};
	$total_amt460300 += $s->{amt460300};
	
	if ($merch_group_code eq 'DS'){
		############DEPT STORE WO JEWELRY ###############
		$worksheet->write($a,7, $grp_o_retail+$grp_c_retail,$headNum);
		$worksheet->write($a,8, $grp_o_margin+$grp_c_margin+$grp_amt501000+$grp_amt432000+$grp_amt433000+$grp_amt458490+$grp_amt503200+$grp_amt503250+$grp_amt503500+$grp_amt506000+$grp_amt503000+$grp_amt507000+$grp_amt999998+$grp_amt505000+$grp_amt504000+$grp_amt434000+$grp_amt458550+$grp_amt460100+$grp_amt460200+$grp_amt460300,$headNum);
			if (($grp_o_retail+$grp_c_retail) le 0){
				$worksheet->write($a,9, "",$headDPct); }
			else{
				$worksheet->write($a,9, ($grp_o_margin+$grp_c_margin+$grp_amt501000+$grp_amt432000+$grp_amt433000+$grp_amt458490+$grp_amt503200+$grp_amt503250+$grp_amt503500+$grp_amt506000+$grp_amt503000+$grp_amt507000+$grp_amt999998+$grp_amt505000+$grp_amt504000+$grp_amt434000+$grp_amt458550+$grp_amt460100+$grp_amt460200+$grp_amt460300)/($grp_o_retail+$grp_c_retail),$headDPct); }
		
		$worksheet->write($a,10, "",$headNum);
		$worksheet->write($a,11, "",$headDPct);
		$worksheet->write($a,12, "",$headNum);
		$worksheet->write($a,13, "",$headDPct);
		
		$worksheet->write($a,14, $grp_o_retail,$headNum);
		$worksheet->write($a,15, $grp_o_margin+$grp_amt501000+$grp_amt432000+$grp_amt433000+$grp_amt458490+$grp_amt503200+$grp_amt503250+$grp_amt503500+$grp_amt506000+$grp_amt503000+$grp_amt507000+$grp_amt999998+$grp_amt505000+$grp_amt504000,$headNum);
			if ($grp_o_retail le 0){
				$worksheet->write($a,16, "",$headDPct); }
			else{
				$worksheet->write($a,16, ($grp_o_margin+$grp_amt501000+$grp_amt432000+$grp_amt433000+$grp_amt458490+$grp_amt503200+$grp_amt503250+$grp_amt503500+$grp_amt506000+$grp_amt503000+$grp_amt507000+$grp_amt999998+$grp_amt505000+$grp_amt504000)/$grp_o_retail,$headDPct); }
		
		$worksheet->write($a,17, "",$headNum);
		$worksheet->write($a,18, "",$headDPct);
		$worksheet->write($a,19, "",$headNum);
		$worksheet->write($a,20, "",$headDPct);			
		
		$worksheet->write($a,21, $grp_o_retail,$headNum);
		$worksheet->write($a,22, $grp_o_margin,$headNum);
			if ($grp_o_retail le 0){
				$worksheet->write($a,23, "",$headDPct); }
			else{
				$worksheet->write($a,23, $grp_o_margin/$grp_o_retail,$headDPct); }
				
		$worksheet->write($a,24, "",$headNum);
		$worksheet->write($a,25, "",$headDPct);
		$worksheet->write($a,26, "",$headNum);
		$worksheet->write($a,27, "",$headDPct);	
		
		$worksheet->write($a,28, $grp_amt501000,$headNum);
		$worksheet->write($a,29, "",$headNum);
		$worksheet->write($a,30, "",$headNum);
			
		$worksheet->write($a,31, $grp_amt503200,$headNum);
		$worksheet->write($a,32, "",$headNum);
		$worksheet->write($a,33, "",$headNum);
		
		$worksheet->write($a,34, $grp_amt503250,$headNum);
		$worksheet->write($a,35, "",$headNum);
		$worksheet->write($a,36, "",$headNum);
		
		$worksheet->write($a,37, $grp_amt503500,$headNum);
		$worksheet->write($a,38, "",$headNum);
		$worksheet->write($a,39, "",$headNum);
		
		$worksheet->write($a,40, $grp_amt506000,$headNum);
		$worksheet->write($a,41, "",$headNum);
		$worksheet->write($a,42, "",$headNum);
			
		$worksheet->write($a,43, $grp_amt503000,$headNum);
		$worksheet->write($a,44, "",$headNum);
		$worksheet->write($a,45, "",$headNum);
					
		$worksheet->write($a,46, $grp_amt507000,$headNum);
		$worksheet->write($a,47, "",$headNum);
		$worksheet->write($a,48, "",$headNum);
		
		$worksheet->write($a,49, $grp_amt999998,$headNum);
		$worksheet->write($a,50, "",$headNum);
		$worksheet->write($a,51, "",$headNum);
		
		$worksheet->write($a,52, $grp_amt505000,$headNum);
		$worksheet->write($a,53, "",$headNum);
		$worksheet->write($a,54, "",$headNum);
					
		$worksheet->write($a,55, $grp_amt432000,$headNum);
		$worksheet->write($a,56, "",$headNum);
		$worksheet->write($a,57, "",$headNum);
		
		$worksheet->write($a,58, $grp_amt433000,$headNum);
		$worksheet->write($a,59, "",$headNum);
		$worksheet->write($a,60, "",$headNum);
		
		$worksheet->write($a,61, $grp_amt458490,$headNum);
		$worksheet->write($a,62, "",$headNum);
		$worksheet->write($a,63, "",$headNum);
		
		$worksheet->write($a,64, $grp_amt504000,$headNum);
		$worksheet->write($a,65, "",$headNum);
		$worksheet->write($a,66, "",$headNum);
		
		$worksheet->write($a,67, $grp_c_retail,$headNum);
		$worksheet->write($a,68, $grp_c_margin+$grp_amt434000+$grp_amt458550+$grp_amt460100+$grp_amt460200+$grp_amt460300,$headNum);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,69, "",$headDPct); }
			else{
				$worksheet->write($a,69, ($grp_c_margin+$grp_amt434000+$grp_amt458550+$grp_amt460100+$grp_amt460200+$grp_amt460300)/$grp_c_retail,$headDPct); }
				
		$worksheet->write($a,70, "",$headNum);
		$worksheet->write($a,71, "",$headDPct);
		$worksheet->write($a,72, "",$headNum);
		$worksheet->write($a,73, "",$headDPct);	
		
		$worksheet->write($a,74, $grp_c_retail,$headNum);
		$worksheet->write($a,75, $grp_c_margin,$headNum);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,76, "",$headDPct); }
			else{
				$worksheet->write($a,76, $grp_c_margin/$grp_c_retail,$headDPct); }
			
		$worksheet->write($a,77, "",$headNum);
		$worksheet->write($a,78, "",$headDPct);
		$worksheet->write($a,79, "",$headNum);
		$worksheet->write($a,80, "",$headDPct);	
				
		$worksheet->write($a,81, $grp_amt434000,$headNum);
		$worksheet->write($a,82, "",$headNum);
		$worksheet->write($a,83, "",$headNum);
				
		$worksheet->write($a,84, $grp_amt458550,$headNum);
		$worksheet->write($a,85, "",$headNum);
		$worksheet->write($a,86, "",$headNum);
		
		$worksheet->write($a,87, $grp_amt460100,$headNum);
		$worksheet->write($a,88, "",$headNum);
		$worksheet->write($a,89, "",$headNum);
			
		$worksheet->write($a,90, $grp_amt460200,$headNum);
		$worksheet->write($a,91, "",$headNum);
		$worksheet->write($a,92, "",$headNum);
				
		$worksheet->write($a,93, $grp_amt460300,$headNum);
		$worksheet->write($a,94, "",$headNum);
		$worksheet->write($a,95, "",$headNum);			
		
		#$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store w/o Jewelry', $headD );
		#$a += 1;
		
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
	}
	
	elsif($merch_group_code eq 'SU'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
	}
	
	elsif($merch_group_code eq 'Z'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Others', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'Others', $border2 );
	}
	
	$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$headNumber);
	$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$headNumber);
		if (($s->{o_retail}+$s->{c_retail}) le 0){
			$worksheet->write($a,9, "",$headPct); }
		else{
			$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$headPct); }
				
	$worksheet->write($a,10, "",$headNumber);
	$worksheet->write($a,11, "",$headPct);
	$worksheet->write($a,12, "",$headNumber);
	$worksheet->write($a,13, "",$headPct);
	
	$worksheet->write($a,14, $s->{o_retail},$headNumber);
	$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$headNumber);
		if ($s->{o_retail} le 0){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$headPct); }
	
	$worksheet->write($a,17, "",$headNumber);
	$worksheet->write($a,18, "",$headPct);
	$worksheet->write($a,19, "",$headNumber);
	$worksheet->write($a,20, "",$headPct);			
	
	$worksheet->write($a,21, $s->{o_retail},$headNumber);
	$worksheet->write($a,22, $s->{o_margin},$headNumber);
		if ($s->{o_retail} le 0){
			$worksheet->write($a,23, "",$headPct); }
		else{
			$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$headPct); }
				
	$worksheet->write($a,24, "",$headNumber);
	$worksheet->write($a,25, "",$headPct);
	$worksheet->write($a,26, "",$headNumber);
	$worksheet->write($a,27, "",$headPct);	
	
	$worksheet->write($a,28, $s->{amt501000},$headNumber);
	$worksheet->write($a,29, "",$headNumber);
	$worksheet->write($a,30, "",$headNumber);
	
	$worksheet->write($a,31, $s->{amt503200},$headNumber);
	$worksheet->write($a,32, "",$headNumber);
	$worksheet->write($a,33, "",$headNumber);
				
	$worksheet->write($a,34, $s->{amt503250},$headNumber);
	$worksheet->write($a,35, "",$headNumber);
	$worksheet->write($a,36, "",$headNumber);
		
	$worksheet->write($a,37, $s->{amt503500},$headNumber);
	$worksheet->write($a,38, "",$headNumber);
	$worksheet->write($a,39, "",$headNumber);
		
	$worksheet->write($a,40, $s->{amt506000},$headNumber);
	$worksheet->write($a,41, "",$headNumber);
	$worksheet->write($a,42, "",$headNumber);
	
	$worksheet->write($a,43, $s->{amt503000},$headNumber);
	$worksheet->write($a,44, "",$headNumber);
	$worksheet->write($a,45, "",$headNumber);
					
	$worksheet->write($a,46, $s->{amt507000},$headNumber);
	$worksheet->write($a,47, "",$headNumber);
	$worksheet->write($a,48, "",$headNumber);
	
	$worksheet->write($a,49, $s->{amt999998},$headNumber);
	$worksheet->write($a,50, "",$headNumber);
	$worksheet->write($a,51, "",$headNumber);
	
	$worksheet->write($a,52, $s->{amt505000},$headNumber);
	$worksheet->write($a,53, "",$headNumber);
	$worksheet->write($a,54, "",$headNumber);
				
	$worksheet->write($a,55, $s->{amt432000},$headNumber);
	$worksheet->write($a,56, "",$headNumber);
	$worksheet->write($a,57, "",$headNumber);
	
	$worksheet->write($a,58, $s->{amt433000},$headNumber);
	$worksheet->write($a,59, "",$headNumber);
	$worksheet->write($a,60, "",$headNumber);
	
	$worksheet->write($a,61, $s->{amt458490},$headNumber);
	$worksheet->write($a,62, "",$headNumber);
	$worksheet->write($a,63, "",$headNumber);
	
	$worksheet->write($a,64, $s->{amt504000},$headNumber);
	$worksheet->write($a,65, "",$headNumber);
	$worksheet->write($a,66, "",$headNumber);
	
	$worksheet->write($a,67, $s->{c_retail},$headNumber);
	$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$headNumber);
		if ($s->{c_retail} le 0){
			$worksheet->write($a,69, "",$headPct); }
		else{
			$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$headPct); }
			
	$worksheet->write($a,70, "",$headNumber);
	$worksheet->write($a,71, "",$headPct);
	$worksheet->write($a,72, "",$headNumber);
	$worksheet->write($a,73, "",$headPct);	
	
	$worksheet->write($a,74, $s->{c_retail},$headNumber);
	$worksheet->write($a,75, $s->{c_margin},$headNumber);
		if ($s->{c_retail} le 0){
			$worksheet->write($a,76, "",$headPct); }
		else{
			$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$headPct); }
		
	$worksheet->write($a,77, "",$headNumber);
	$worksheet->write($a,78, "",$headPct);
	$worksheet->write($a,79, "",$headNumber);
	$worksheet->write($a,80, "",$headPct);	
			
	$worksheet->write($a,81, $s->{amt434000},$headNumber);
	$worksheet->write($a,82, "",$headNumber);
	$worksheet->write($a,83, "",$headNumber);
			
	$worksheet->write($a,84, $s->{amt458550},$headNumber);
	$worksheet->write($a,85, "",$headNumber);
	$worksheet->write($a,86, "",$headNumber);
	
	$worksheet->write($a,87, $s->{amt460100},$headNumber);
	$worksheet->write($a,88, "",$headNumber);
	$worksheet->write($a,89, "",$headNumber);
		
	$worksheet->write($a,90, $s->{amt460200},$headNumber);
	$worksheet->write($a,91, "",$headNumber);
	$worksheet->write($a,92, "",$headNumber);
			
	$worksheet->write($a,93, $s->{amt460300},$headNumber);
	$worksheet->write($a,94, "",$headNumber);
	$worksheet->write($a,95, "",$headNumber);		
	
	$a++; #INCREMENT VARIABLE a
}

	$worksheet->write($a,7, $total_o_retail+$total_c_retail,$headNumber);
	$worksheet->write($a,8, $total_o_margin+$total_c_margin+$total_amt501000+$total_amt432000+$total_amt433000+$total_amt458490+$total_amt503200+$total_amt503250+$total_amt503500+$total_amt506000+$total_amt503000+$total_amt507000+$total_amt999998+$total_amt505000+$total_amt504000+$total_amt434000+$total_amt458550+$total_amt460100+$total_amt460200+$total_amt460300,$headNumber);
			if (($total_o_retail+$total_c_retail) le 0){
				$worksheet->write($a,9, "",$headPct); }
			else{
				$worksheet->write($a,9, ($total_o_margin+$total_c_margin+$total_amt501000+$total_amt432000+$total_amt433000+$total_amt458490+$total_amt503200+$total_amt503250+$total_amt503500+$total_amt506000+$total_amt503000+$total_amt507000+$total_amt999998+$total_amt505000+$total_amt504000+$total_amt434000+$total_amt458550+$total_amt460100+$total_amt460200+$total_amt460300)/($total_o_retail+$total_c_retail),$headPct); }
		
	$worksheet->write($a,10, "",$headNumber);
	$worksheet->write($a,11, "",$headPct);
	$worksheet->write($a,12, "",$headNumber);
	$worksheet->write($a,13, "",$headPct);
	
	$worksheet->write($a,14, $total_o_retail,$headNumber);
	$worksheet->write($a,15, $total_o_margin+$total_amt501000+$total_amt432000+$total_amt433000+$total_amt458490+$total_amt503200+$total_amt503250+$total_amt503500+$total_amt506000+$total_amt503000+$total_amt507000+$total_amt999998+$total_amt505000+$total_amt504000,$headNumber);
		if ($total_o_retail le 0){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($total_o_margin+$total_amt501000+$total_amt432000+$total_amt433000+$total_amt458490+$total_amt503200+$total_amt503250+$total_amt503500+$total_amt506000+$total_amt503000+$total_amt507000+$total_amt999998+$total_amt505000+$total_amt504000)/$total_o_retail,$headPct); }
	
	$worksheet->write($a,17, "",$headNumber);
	$worksheet->write($a,18, "",$headPct);
	$worksheet->write($a,19, "",$headNumber);
	$worksheet->write($a,20, "",$headPct);			
	
	$worksheet->write($a,21, $total_o_retail,$headNumber);
	$worksheet->write($a,22, $total_o_margin,$headNumber);
		if ($total_o_retail le 0){
			$worksheet->write($a,23, "",$headPct); }
		else{
			$worksheet->write($a,23, $total_o_margin/$total_o_retail,$headPct); }
			
	$worksheet->write($a,24, "",$headNumber);
	$worksheet->write($a,25, "",$headPct);
	$worksheet->write($a,26, "",$headNumber);
	$worksheet->write($a,27, "",$headPct);	
	
	$worksheet->write($a,28, $total_amt501000,$headNumber);
	$worksheet->write($a,29, "",$headNumber);
	$worksheet->write($a,30, "",$headNumber);
		
	$worksheet->write($a,31, $total_amt503200,$headNumber);
	$worksheet->write($a,32, "",$headNumber);
	$worksheet->write($a,33, "",$headNumber);
	
	$worksheet->write($a,34, $total_amt503250,$headNumber);
	$worksheet->write($a,35, "",$headNumber);
	$worksheet->write($a,36, "",$headNumber);
	
	$worksheet->write($a,37, $total_amt503500,$headNumber);
	$worksheet->write($a,38, "",$headNumber);
	$worksheet->write($a,39, "",$headNumber);
	
	$worksheet->write($a,40, $total_amt506000,$headNumber);
	$worksheet->write($a,41, "",$headNumber);
	$worksheet->write($a,42, "",$headNumber);
		
	$worksheet->write($a,43, $total_amt503000,$headNumber);
	$worksheet->write($a,44, "",$headNumber);
	$worksheet->write($a,45, "",$headNumber);
				
	$worksheet->write($a,46, $total_amt507000,$headNumber);
	$worksheet->write($a,47, "",$headNumber);
	$worksheet->write($a,48, "",$headNumber);
	
	$worksheet->write($a,49, $total_amt999998,$headNumber);
	$worksheet->write($a,50, "",$headNumber);
	$worksheet->write($a,51, "",$headNumber);
	
	$worksheet->write($a,52, $total_amt505000,$headNumber);
	$worksheet->write($a,53, "",$headNumber);
	$worksheet->write($a,54, "",$headNumber);
				
	$worksheet->write($a,55, $total_amt432000,$headNumber);
	$worksheet->write($a,56, "",$headNumber);
	$worksheet->write($a,57, "",$headNumber);
	
	$worksheet->write($a,58, $total_amt433000,$headNumber);
	$worksheet->write($a,59, "",$headNumber);
	$worksheet->write($a,60, "",$headNumber);
	
	$worksheet->write($a,61, $total_amt458490,$headNumber);
	$worksheet->write($a,62, "",$headNumber);
	$worksheet->write($a,63, "",$headNumber);
	
	$worksheet->write($a,64, $total_amt504000,$headNumber);
	$worksheet->write($a,65, "",$headNumber);
	$worksheet->write($a,66, "",$headNumber);
	
	$worksheet->write($a,67, $total_c_retail,$headNumber);
	$worksheet->write($a,68, $total_c_margin+$total_amt434000+$total_amt458550+$total_amt460100+$total_amt460200+$total_amt460300,$headNumber);
		if ($s->{c_retail} le 0){
			$worksheet->write($a,69, "",$headPct); }
		else{
			$worksheet->write($a,69, ($total_c_margin+$total_amt434000+$total_amt458550+$total_amt460100+$total_amt460200+$total_amt460300)/$total_c_retail,$headPct); }
			
	$worksheet->write($a,70, "",$headNumber);
	$worksheet->write($a,71, "",$headPct);
	$worksheet->write($a,72, "",$headNumber);
	$worksheet->write($a,73, "",$headPct);	
	
	$worksheet->write($a,74, $total_c_retail,$headNumber);
	$worksheet->write($a,75, $total_c_margin,$headNumber);
		if ($s->{c_retail} le 0){
			$worksheet->write($a,76, "",$headPct); }
		else{
			$worksheet->write($a,76, $total_c_margin/$total_c_retail,$headPct); }
		
	$worksheet->write($a,77, "",$headNumber);
	$worksheet->write($a,78, "",$headPct);
	$worksheet->write($a,79, "",$headNumber);
	$worksheet->write($a,80, "",$headPct);	
			
	$worksheet->write($a,81, $total_amt434000,$headNumber);
	$worksheet->write($a,82, "",$headNumber);
	$worksheet->write($a,83, "",$headNumber);
			
	$worksheet->write($a,84, $total_amt458550,$headNumber);
	$worksheet->write($a,85, "",$headNumber);
	$worksheet->write($a,86, "",$headNumber);
	
	$worksheet->write($a,87, $total_amt460100,$headNumber);
	$worksheet->write($a,88, "",$headNumber);
	$worksheet->write($a,89, "",$headNumber);
		
	$worksheet->write($a,90, $total_amt460200,$headNumber);
	$worksheet->write($a,91, "",$headNumber);
	$worksheet->write($a,92, "",$headNumber);
			
	$worksheet->write($a,93, $total_amt460300,$headNumber);
	$worksheet->write($a,94, "",$headNumber);
	$worksheet->write($a,95, "",$headNumber);			
	
	$worksheet->write($loc, 2, $loc_desc, $bold);
	$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}

sub query_dept_store {

$table = 'consolidated_margin_x.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls1 = $dbh_csv->prepare (qq{SELECT store_code, store_description, merch_group_code_rev, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
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
	
	$sls2 = $dbh_csv->prepare (qq{SELECT group_code, group_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
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
				
		$sls3 = $dbh_csv->prepare (qq{SELECT division, division_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
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
			
			$sls4 = $dbh_csv->prepare (qq{SELECT department_code, department_desc, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
										 FROM $table 
										 WHERE group_code = '$group_code' and merch_group_code_rev = '$merch_group_code' and division = '$division' and store_code = '$store'
										 GROUP BY department_code, department_desc 
										 ORDER BY department_code
										 });
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){
				
				$worksheet->write($a,5, $s->{department_code},$desc);
				$worksheet->write($a,6, $s->{department_desc},$desc);
				
				$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1);
				$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if (($s->{o_retail}+$s->{c_retail}) le 0){
						$worksheet->write($a,9, "",$subt); }
					else{
						$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
				$worksheet->write($a,10, "",$border1);
				$worksheet->write($a,11, "",$subt);
				$worksheet->write($a,12, "",$border1);
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{o_retail},$border1);
				$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1);
					if ($s->{o_retail} le 0){
						$worksheet->write($a,16, "",$subt); }
					else{
						$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
				
				$worksheet->write($a,17, "",$border1);
				$worksheet->write($a,18, "",$subt);
				$worksheet->write($a,19, "",$border1);
				$worksheet->write($a,20, "",$subt);			
				
				$worksheet->write($a,21, $s->{o_retail},$border1);
				$worksheet->write($a,22, $s->{o_margin},$border1);
					if ($s->{o_retail} le 0){
						$worksheet->write($a,23, "",$subt); }
					else{
						$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
						
				$worksheet->write($a,24, "",$border1);
				$worksheet->write($a,25, "",$subt);
				$worksheet->write($a,26, "",$border1);
				$worksheet->write($a,27, "",$subt);	
				
				$worksheet->write($a,28, $s->{amt501000},$border1);
				$worksheet->write($a,29, "",$border1);
				$worksheet->write($a,30, "",$border1);
				
				$worksheet->write($a,31, $s->{amt503200},$border1);
				$worksheet->write($a,32, "",$border1);
				$worksheet->write($a,33, "",$border1);
							
				$worksheet->write($a,34, $s->{amt503250},$border1);
				$worksheet->write($a,35, "",$border1);
				$worksheet->write($a,36, "",$border1);
					
				$worksheet->write($a,37, $s->{amt503500},$border1);
				$worksheet->write($a,38, "",$border1);
				$worksheet->write($a,39, "",$border1);
					
				$worksheet->write($a,40, $s->{amt506000},$border1);
				$worksheet->write($a,41, "",$border1);
				$worksheet->write($a,42, "",$border1);
				
				$worksheet->write($a,43, $s->{amt503000},$border1);
				$worksheet->write($a,44, "",$border1);
				$worksheet->write($a,45, "",$border1);
							
				$worksheet->write($a,46, $s->{amt507000},$border1);
				$worksheet->write($a,47, "",$border1);
				$worksheet->write($a,48, "",$border1);
				
				$worksheet->write($a,49, $s->{amt999998},$border1);
				$worksheet->write($a,50, "",$border1);
				$worksheet->write($a,51, "",$border1);
				
				$worksheet->write($a,52, $s->{amt505000},$border1);
				$worksheet->write($a,53, "",$border1);
				$worksheet->write($a,54, "",$border1);
							
				$worksheet->write($a,55, $s->{amt432000},$border1);
				$worksheet->write($a,56, "",$border1);
				$worksheet->write($a,57, "",$border1);
				
				$worksheet->write($a,58, $s->{amt433000},$border1);
				$worksheet->write($a,59, "",$border1);
				$worksheet->write($a,60, "",$border1);
				
				$worksheet->write($a,61, $s->{amt458490},$border1);
				$worksheet->write($a,62, "",$border1);
				$worksheet->write($a,63, "",$border1);
				
				$worksheet->write($a,64, $s->{amt504000},$border1);
				$worksheet->write($a,65, "",$border1);
				$worksheet->write($a,66, "",$border1);
				
				$worksheet->write($a,67, $s->{c_retail},$border1);
				$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,69, "",$subt); }
					else{
						$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
						
				$worksheet->write($a,70, "",$border1);
				$worksheet->write($a,71, "",$subt);
				$worksheet->write($a,72, "",$border1);
				$worksheet->write($a,73, "",$subt);	
				
				$worksheet->write($a,74, $s->{c_retail},$border1);
				$worksheet->write($a,75, $s->{c_margin},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,76, "",$subt); }
					else{
						$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
					
				$worksheet->write($a,77, "",$border1);
				$worksheet->write($a,78, "",$subt);
				$worksheet->write($a,79, "",$border1);
				$worksheet->write($a,80, "",$subt);	
						
				$worksheet->write($a,81, $s->{amt434000},$border1);
				$worksheet->write($a,82, "",$border1);
				$worksheet->write($a,83, "",$border1);
						
				$worksheet->write($a,84, $s->{amt458550},$border1);
				$worksheet->write($a,85, "",$border1);
				$worksheet->write($a,86, "",$border1);
				
				$worksheet->write($a,87, $s->{amt460100},$border1);
				$worksheet->write($a,88, "",$border1);
				$worksheet->write($a,89, "",$border1);
					
				$worksheet->write($a,90, $s->{amt460200},$border1);
				$worksheet->write($a,91, "",$border1);
				$worksheet->write($a,92, "",$border1);
						
				$worksheet->write($a,93, $s->{amt460300},$border1);
				$worksheet->write($a,94, "",$border1);
				$worksheet->write($a,95, "",$border1);			
				
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
			$grp_amt501000 += $s->{amt501000};
			$grp_amt503200 += $s->{amt503200};
			$grp_amt503250 += $s->{amt503250}; 
			$grp_amt503500 += $s->{amt503500};
			$grp_amt506000 += $s->{amt506000};
			$grp_amt503000 += $s->{amt503000};
			$grp_amt507000 += $s->{amt507000};
			$grp_amt999998 += $s->{amt999998};
			$grp_amt504000 += $s->{amt504000};
			$grp_amt505000 += $s->{amt505000};
			$grp_amt432000 += $s->{amt432000};
			$grp_amt433000 += $s->{amt433000}; 
			$grp_amt458490 += $s->{amt458490};
			$grp_amt434000 += $s->{amt434000};
			$grp_amt458550 += $s->{amt458550};
			$grp_amt460100 += $s->{amt460100};
			$grp_amt460200 += $s->{amt460200};
			$grp_amt460300 += $s->{amt460300};
		}
		
		$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$bodyNum);
		$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$bodyNum);
			if (($s->{o_retail}+$s->{c_retail}) le 0){
				$worksheet->write($a,9, "",$bodyPct); }
			else{
				$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$bodyPct); }
				
		$worksheet->write($a,10, "",$bodyNum);
		$worksheet->write($a,11, "",$bodyPct);
		$worksheet->write($a,12, "",$bodyNum);
		$worksheet->write($a,13, "",$bodyPct);
				
		$worksheet->write($a,14, $s->{o_retail},$bodyNum);
		$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$bodyNum);
			if ($s->{o_retail} le 0){
				$worksheet->write($a,16, "",$bodyPct); }
			else{
				$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$bodyPct); }
				
		$worksheet->write($a,17, "",$bodyNum);
		$worksheet->write($a,18, "",$bodyPct);
		$worksheet->write($a,19, "",$bodyNum);
		$worksheet->write($a,20, "",$bodyPct);			
				
		$worksheet->write($a,21, $s->{o_retail},$bodyNum);
		$worksheet->write($a,22, $s->{o_margin},$bodyNum);
			if ($s->{o_retail} le 0){
				$worksheet->write($a,23, "",$bodyPct); }
			else{
				$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$bodyPct); }
						
		$worksheet->write($a,24, "",$bodyNum);
		$worksheet->write($a,25, "",$bodyPct);
		$worksheet->write($a,26, "",$bodyNum);
		$worksheet->write($a,27, "",$bodyPct);	
		
		$worksheet->write($a,28, $s->{amt501000},$bodyNum);
		$worksheet->write($a,29, "",$bodyNum);
		$worksheet->write($a,30, "",$bodyNum);
		
		$worksheet->write($a,31, $s->{amt503200},$bodyNum);
		$worksheet->write($a,32, "",$bodyNum);
		$worksheet->write($a,33, "",$bodyNum);
					
		$worksheet->write($a,34, $s->{amt503250},$bodyNum);
		$worksheet->write($a,35, "",$bodyNum);
		$worksheet->write($a,36, "",$bodyNum);
			
		$worksheet->write($a,37, $s->{amt503500},$bodyNum);
		$worksheet->write($a,38, "",$bodyNum);
		$worksheet->write($a,39, "",$bodyNum);
			
		$worksheet->write($a,40, $s->{amt506000},$bodyNum);
		$worksheet->write($a,41, "",$bodyNum);
		$worksheet->write($a,42, "",$bodyNum);
		
		$worksheet->write($a,43, $s->{amt503000},$bodyNum);
		$worksheet->write($a,44, "",$bodyNum);
		$worksheet->write($a,45, "",$bodyNum);
							
		$worksheet->write($a,46, $s->{amt507000},$bodyNum);
		$worksheet->write($a,47, "",$bodyNum);
		$worksheet->write($a,48, "",$bodyNum);
		
		$worksheet->write($a,49, $s->{amt999998},$bodyNum);
		$worksheet->write($a,50, "",$bodyNum);
		$worksheet->write($a,51, "",$bodyNum);
		
		$worksheet->write($a,52, $s->{amt505000},$bodyNum);
		$worksheet->write($a,53, "",$bodyNum);
		$worksheet->write($a,54, "",$bodyNum);
					
		$worksheet->write($a,55, $s->{amt432000},$bodyNum);
		$worksheet->write($a,56, "",$bodyNum);
		$worksheet->write($a,57, "",$bodyNum);
		
		$worksheet->write($a,58, $s->{amt433000},$bodyNum);
		$worksheet->write($a,59, "",$bodyNum);
		$worksheet->write($a,60, "",$bodyNum);
		
		$worksheet->write($a,61, $s->{amt458490},$bodyNum);
		$worksheet->write($a,62, "",$bodyNum);
		$worksheet->write($a,63, "",$bodyNum);
		
		$worksheet->write($a,64, $s->{amt504000},$bodyNum);
		$worksheet->write($a,65, "",$bodyNum);
		$worksheet->write($a,66, "",$bodyNum);
		
		$worksheet->write($a,67, $s->{c_retail},$bodyNum);
		$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$bodyNum);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,69, "",$bodyPct); }
			else{
				$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$bodyPct); }
				
		$worksheet->write($a,70, "",$bodyNum);
		$worksheet->write($a,71, "",$bodyPct);
		$worksheet->write($a,72, "",$bodyNum);
		$worksheet->write($a,73, "",$bodyPct);	
		
		$worksheet->write($a,74, $s->{c_retail},$bodyNum);
		$worksheet->write($a,75, $s->{c_margin},$bodyNum);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,76, "",$bodyPct); }
			else{
				$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$bodyPct); }
			
		$worksheet->write($a,77, "",$bodyNum);
		$worksheet->write($a,78, "",$bodyPct);
		$worksheet->write($a,79, "",$bodyNum);
		$worksheet->write($a,80, "",$bodyPct);	
				
		$worksheet->write($a,81, $s->{amt434000},$bodyNum);
		$worksheet->write($a,82, "",$bodyNum);
		$worksheet->write($a,83, "",$bodyNum);
				
		$worksheet->write($a,84, $s->{amt458550},$bodyNum);
		$worksheet->write($a,85, "",$bodyNum);
		$worksheet->write($a,86, "",$bodyNum);
		
		$worksheet->write($a,87, $s->{amt460100},$bodyNum);
		$worksheet->write($a,88, "",$bodyNum);
		$worksheet->write($a,89, "",$bodyNum);
			
		$worksheet->write($a,90, $s->{amt460200},$bodyNum);
		$worksheet->write($a,91, "",$bodyNum);
		$worksheet->write($a,92, "",$bodyNum);
				
		$worksheet->write($a,93, $s->{amt460300},$bodyNum);
		$worksheet->write($a,94, "",$bodyNum);
		$worksheet->write($a,95, "",$bodyNum);					

		$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

		$a++; #INCREMENT VARIABLE a
	}
	
	$total_o_retail += $s->{o_retail};
	$total_o_margin += $s->{o_margin};
	$total_c_retail += $s->{c_retail};
	$total_c_margin += $s->{c_margin};
	$total_amt501000 += $s->{amt501000};
	$total_amt503200 += $s->{amt503200};
	$total_amt503250 += $s->{amt503250};
	$total_amt503500 += $s->{amt503500};
	$total_amt506000 += $s->{amt506000};
	$total_amt503000 += $s->{amt503000};
	$total_amt507000 += $s->{amt507000};
	$total_amt999998 += $s->{amt999998};
	$total_amt504000 += $s->{amt504000};
	$total_amt505000 += $s->{amt505000};
	$total_amt432000 += $s->{amt432000};
	$total_amt433000 += $s->{amt433000};
	$total_amt458490 += $s->{amt458490};
	$total_amt434000 += $s->{amt434000};
	$total_amt458550 += $s->{amt458550};
	$total_amt460100 += $s->{amt460100};
	$total_amt460200 += $s->{amt460200};
	$total_amt460300 += $s->{amt460300};
	
	if ($merch_group_code eq 'DS'){
		############DEPT STORE WO JEWELRY ###############
		$worksheet->write($a,7, $grp_o_retail+$grp_c_retail,$headNum);
		$worksheet->write($a,8, $grp_o_margin+$grp_c_margin+$grp_amt501000+$grp_amt432000+$grp_amt433000+$grp_amt458490+$grp_amt503200+$grp_amt503250+$grp_amt503500+$grp_amt506000+$grp_amt503000+$grp_amt507000+$grp_amt999998+$grp_amt505000+$grp_amt504000+$grp_amt434000+$grp_amt458550+$grp_amt460100+$grp_amt460200+$grp_amt460300,$headNum);
			if (($grp_o_retail+$grp_c_retail) le 0){
				$worksheet->write($a,9, "",$headDPct); }
			else{
				$worksheet->write($a,9, ($grp_o_margin+$grp_c_margin+$grp_amt501000+$grp_amt432000+$grp_amt433000+$grp_amt458490+$grp_amt503200+$grp_amt503250+$grp_amt503500+$grp_amt506000+$grp_amt503000+$grp_amt507000+$grp_amt999998+$grp_amt505000+$grp_amt504000+$grp_amt434000+$grp_amt458550+$grp_amt460100+$grp_amt460200+$grp_amt460300)/($grp_o_retail+$grp_c_retail),$headDPct); }
		
		$worksheet->write($a,10, "",$headNum);
		$worksheet->write($a,11, "",$headDPct);
		$worksheet->write($a,12, "",$headNum);
		$worksheet->write($a,13, "",$headDPct);
		
		$worksheet->write($a,14, $grp_o_retail,$headNum);
		$worksheet->write($a,15, $grp_o_margin+$grp_amt501000+$grp_amt432000+$grp_amt433000+$grp_amt458490+$grp_amt503200+$grp_amt503250+$grp_amt503500+$grp_amt506000+$grp_amt503000+$grp_amt507000+$grp_amt999998+$grp_amt505000+$grp_amt504000,$headNum);
			if ($grp_o_retail le 0){
				$worksheet->write($a,16, "",$headDPct); }
			else{
				$worksheet->write($a,16, ($grp_o_margin+$grp_amt501000+$grp_amt432000+$grp_amt433000+$grp_amt458490+$grp_amt503200+$grp_amt503250+$grp_amt503500+$grp_amt506000+$grp_amt503000+$grp_amt507000+$grp_amt999998+$grp_amt505000+$grp_amt504000)/$grp_o_retail,$headDPct); }
		
		$worksheet->write($a,17, "",$headNum);
		$worksheet->write($a,18, "",$headDPct);
		$worksheet->write($a,19, "",$headNum);
		$worksheet->write($a,20, "",$headDPct);			
		
		$worksheet->write($a,21, $grp_o_retail,$headNum);
		$worksheet->write($a,22, $grp_o_margin,$headNum);
			if ($grp_o_retail le 0){
				$worksheet->write($a,23, "",$headDPct); }
			else{
				$worksheet->write($a,23, $grp_o_margin/$grp_o_retail,$headDPct); }
				
		$worksheet->write($a,24, "",$headNum);
		$worksheet->write($a,25, "",$headDPct);
		$worksheet->write($a,26, "",$headNum);
		$worksheet->write($a,27, "",$headDPct);	
		
		$worksheet->write($a,28, $grp_amt501000,$headNum);
		$worksheet->write($a,29, "",$headNum);
		$worksheet->write($a,30, "",$headNum);
		
		$worksheet->write($a,31, $grp_amt503200,$headNum);
		$worksheet->write($a,32, "",$headNum);
		$worksheet->write($a,33, "",$headNum);
					
		$worksheet->write($a,34, $grp_amt503250,$headNum);
		$worksheet->write($a,35, "",$headNum);
		$worksheet->write($a,36, "",$headNum);
			
		$worksheet->write($a,37, $grp_amt503500,$headNum);
		$worksheet->write($a,38, "",$headNum);
		$worksheet->write($a,39, "",$headNum);
			
		$worksheet->write($a,40, $grp_amt506000,$headNum);
		$worksheet->write($a,41, "",$headNum);
		$worksheet->write($a,42, "",$headNum);
		
		$worksheet->write($a,43, $grp_amt503000,$headNum);
		$worksheet->write($a,44, "",$headNum);
		$worksheet->write($a,45, "",$headNum);
					
		$worksheet->write($a,46, $grp_amt507000,$headNum);
		$worksheet->write($a,47, "",$headNum);
		$worksheet->write($a,48, "",$headNum);
		
		$worksheet->write($a,49, $grp_amt999998,$headNum);
		$worksheet->write($a,50, "",$headNum);
		$worksheet->write($a,51, "",$headNum);
		
		$worksheet->write($a,52, $grp_amt505000,$headNum);
		$worksheet->write($a,53, "",$headNum);
		$worksheet->write($a,54, "",$headNum);
					
		$worksheet->write($a,55, $grp_amt432000,$headNum);
		$worksheet->write($a,56, "",$headNum);
		$worksheet->write($a,57, "",$headNum);
		
		$worksheet->write($a,58, $grp_amt433000,$headNum);
		$worksheet->write($a,59, "",$headNum);
		$worksheet->write($a,60, "",$headNum);
		
		$worksheet->write($a,61, $grp_amt458490,$headNum);
		$worksheet->write($a,62, "",$headNum);
		$worksheet->write($a,63, "",$headNum);
		
		$worksheet->write($a,64, $grp_amt504000,$headNum);
		$worksheet->write($a,65, "",$headNum);
		$worksheet->write($a,66, "",$headNum);
		
		$worksheet->write($a,67, $grp_c_retail,$headNum);
		$worksheet->write($a,68, $grp_c_margin+$grp_amt434000+$grp_amt458550+$grp_amt460100+$grp_amt460200+$grp_amt460300,$headNum);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,69, "",$headDPct); }
			else{
				$worksheet->write($a,69, ($grp_c_margin+$grp_amt434000+$grp_amt458550+$grp_amt460100+$grp_amt460200+$grp_amt460300)/$grp_c_retail,$headDPct); }
				
		$worksheet->write($a,70, "",$headNum);
		$worksheet->write($a,71, "",$headDPct);
		$worksheet->write($a,72, "",$headNum);
		$worksheet->write($a,73, "",$headDPct);	
		
		$worksheet->write($a,74, $grp_c_retail,$headNum);
		$worksheet->write($a,75, $grp_c_margin,$headNum);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,76, "",$headDPct); }
			else{
				$worksheet->write($a,76, $grp_c_margin/$grp_c_retail,$headDPct); }
			
		$worksheet->write($a,77, "",$headNum);
		$worksheet->write($a,78, "",$headDPct);
		$worksheet->write($a,79, "",$headNum);
		$worksheet->write($a,80, "",$headDPct);	
				
		$worksheet->write($a,81, $grp_amt434000,$headNum);
		$worksheet->write($a,82, "",$headNum);
		$worksheet->write($a,83, "",$headNum);
				
		$worksheet->write($a,84, $grp_amt458550,$headNum);
		$worksheet->write($a,85, "",$headNum);
		$worksheet->write($a,86, "",$headNum);
		
		$worksheet->write($a,87, $grp_amt460100,$headNum);
		$worksheet->write($a,88, "",$headNum);
		$worksheet->write($a,89, "",$headNum);
			
		$worksheet->write($a,90, $grp_amt460200,$headNum);
		$worksheet->write($a,91, "",$headNum);
		$worksheet->write($a,92, "",$headNum);
				
		$worksheet->write($a,93, $grp_amt460300,$headNum);
		$worksheet->write($a,94, "",$headNum);
		$worksheet->write($a,95, "",$headNum);		
		
		#$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store w/o Jewelry', $headD );
		#$a += 1;
		
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
	}
	
	elsif($merch_group_code eq 'SU'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
	}
	
	elsif($merch_group_code eq 'Z'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Others', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'Others', $border2 );
	}
	
	$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$headNumber);
	$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$headNumber);
		if (($s->{o_retail}+$s->{c_retail}) le 0){
			$worksheet->write($a,9, "",$headPct); }
		else{
			$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$headPct); }
				
	$worksheet->write($a,10, "",$headNumber);
	$worksheet->write($a,11, "",$headPct);
	$worksheet->write($a,12, "",$headNumber);
	$worksheet->write($a,13, "",$headPct);
	
	$worksheet->write($a,14, $s->{o_retail},$headNumber);
	$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$headNumber);
		if ($s->{o_retail} le 0){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$headPct); }
	
	$worksheet->write($a,17, "",$headNumber);
	$worksheet->write($a,18, "",$headPct);
	$worksheet->write($a,19, "",$headNumber);
	$worksheet->write($a,20, "",$headPct);			
	
	$worksheet->write($a,21, $s->{o_retail},$headNumber);
	$worksheet->write($a,22, $s->{o_margin},$headNumber);
		if ($s->{o_retail} le 0){
			$worksheet->write($a,23, "",$headPct); }
		else{
			$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$headPct); }
				
	$worksheet->write($a,24, "",$headNumber);
	$worksheet->write($a,25, "",$headPct);
	$worksheet->write($a,26, "",$headNumber);
	$worksheet->write($a,27, "",$headPct);	
	
	$worksheet->write($a,28, $s->{amt501000},$headNumber);
	$worksheet->write($a,29, "",$headNumber);
	$worksheet->write($a,30, "",$headNumber);
	
	$worksheet->write($a,31, $s->{amt503200},$headNumber);
	$worksheet->write($a,32, "",$headNumber);
	$worksheet->write($a,33, "",$headNumber);
				
	$worksheet->write($a,34, $s->{amt503250},$headNumber);
	$worksheet->write($a,35, "",$headNumber);
	$worksheet->write($a,36, "",$headNumber);
		
	$worksheet->write($a,37, $s->{amt503500},$headNumber);
	$worksheet->write($a,38, "",$headNumber);
	$worksheet->write($a,39, "",$headNumber);
		
	$worksheet->write($a,40, $s->{amt506000},$headNumber);
	$worksheet->write($a,41, "",$headNumber);
	$worksheet->write($a,42, "",$headNumber);
	
	$worksheet->write($a,43, $s->{amt503000},$headNumber);
	$worksheet->write($a,44, "",$headNumber);
	$worksheet->write($a,45, "",$headNumber);
					
	$worksheet->write($a,46, $s->{amt507000},$headNumber);
	$worksheet->write($a,47, "",$headNumber);
	$worksheet->write($a,48, "",$headNumber);
	
	$worksheet->write($a,49, $s->{amt999998},$headNumber);
	$worksheet->write($a,50, "",$headNumber);
	$worksheet->write($a,51, "",$headNumber);
	
	$worksheet->write($a,52, $s->{amt505000},$headNumber);
	$worksheet->write($a,53, "",$headNumber);
	$worksheet->write($a,54, "",$headNumber);
				
	$worksheet->write($a,55, $s->{amt432000},$headNumber);
	$worksheet->write($a,56, "",$headNumber);
	$worksheet->write($a,57, "",$headNumber);
	
	$worksheet->write($a,58, $s->{amt433000},$headNumber);
	$worksheet->write($a,59, "",$headNumber);
	$worksheet->write($a,60, "",$headNumber);
	
	$worksheet->write($a,61, $s->{amt458490},$headNumber);
	$worksheet->write($a,62, "",$headNumber);
	$worksheet->write($a,63, "",$headNumber);
	
	$worksheet->write($a,64, $s->{amt504000},$headNumber);
	$worksheet->write($a,65, "",$headNumber);
	$worksheet->write($a,66, "",$headNumber);
	
	$worksheet->write($a,67, $s->{c_retail},$headNumber);
	$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$headNumber);
		if ($s->{c_retail} le 0){
			$worksheet->write($a,69, "",$headPct); }
		else{
			$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$headPct); }
			
	$worksheet->write($a,70, "",$headNumber);
	$worksheet->write($a,71, "",$headPct);
	$worksheet->write($a,72, "",$headNumber);
	$worksheet->write($a,73, "",$headPct);	
	
	$worksheet->write($a,74, $s->{c_retail},$headNumber);
	$worksheet->write($a,75, $s->{c_margin},$headNumber);
		if ($s->{c_retail} le 0){
			$worksheet->write($a,76, "",$headPct); }
		else{
			$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$headPct); }
		
	$worksheet->write($a,77, "",$headNumber);
	$worksheet->write($a,78, "",$headPct);
	$worksheet->write($a,79, "",$headNumber);
	$worksheet->write($a,80, "",$headPct);	
			
	$worksheet->write($a,81, $s->{amt434000},$headNumber);
	$worksheet->write($a,82, "",$headNumber);
	$worksheet->write($a,83, "",$headNumber);
			
	$worksheet->write($a,84, $s->{amt458550},$headNumber);
	$worksheet->write($a,85, "",$headNumber);
	$worksheet->write($a,86, "",$headNumber);
	
	$worksheet->write($a,87, $s->{amt460100},$headNumber);
	$worksheet->write($a,88, "",$headNumber);
	$worksheet->write($a,89, "",$headNumber);
		
	$worksheet->write($a,90, $s->{amt460200},$headNumber);
	$worksheet->write($a,91, "",$headNumber);
	$worksheet->write($a,92, "",$headNumber);
			
	$worksheet->write($a,93, $s->{amt460300},$headNumber);
	$worksheet->write($a,94, "",$headNumber);
	$worksheet->write($a,95, "",$headNumber);		
	
	$a++; #INCREMENT VARIABLE a
}

	$worksheet->write($a,7, $total_o_retail+$total_c_retail,$headNumber);
	$worksheet->write($a,8, $total_o_margin+$total_c_margin+$total_amt501000+$total_amt432000+$total_amt433000+$total_amt458490+$total_amt503200+$total_amt503250+$total_amt503500+$total_amt506000+$total_amt503000+$total_amt507000+$total_amt999998+$total_amt505000+$total_amt504000+$total_amt434000+$total_amt458550+$total_amt460100+$total_amt460200+$total_amt460300,$headNumber);
			if (($total_o_retail+$total_c_retail) le 0){
				$worksheet->write($a,9, "",$headPct); }
			else{
				$worksheet->write($a,9, ($total_o_margin+$total_c_margin+$total_amt501000+$total_amt432000+$total_amt433000+$total_amt458490+$total_amt503200+$total_amt503250+$total_amt503500+$total_amt506000+$total_amt503000+$total_amt507000+$total_amt999998+$total_amt505000+$total_amt504000+$total_amt434000+$total_amt458550+$total_amt460100+$total_amt460200+$total_amt460300)/($total_o_retail+$total_c_retail),$headPct); }
		
	$worksheet->write($a,10, "",$headNumber);
	$worksheet->write($a,11, "",$headPct);
	$worksheet->write($a,12, "",$headNumber);
	$worksheet->write($a,13, "",$headPct);
	
	$worksheet->write($a,14, $total_o_retail,$headNumber);
	$worksheet->write($a,15, $total_o_margin+$total_amt501000+$total_amt432000+$total_amt433000+$total_amt458490+$total_amt503200+$total_amt503250+$total_amt503500+$total_amt506000+$total_amt503000+$total_amt507000+$total_amt999998+$total_amt505000+$total_amt504000,$headNumber);
		if ($total_o_retail le 0){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($total_o_margin+$total_amt501000+$total_amt432000+$total_amt433000+$total_amt458490+$total_amt503200+$total_amt503250+$total_amt503500+$total_amt506000+$total_amt503000+$total_amt507000+$total_amt999998+$total_amt505000+$total_amt504000)/$total_o_retail,$headPct); }
	
	$worksheet->write($a,17, "",$headNumber);
	$worksheet->write($a,18, "",$headPct);
	$worksheet->write($a,19, "",$headNumber);
	$worksheet->write($a,20, "",$headPct);			
	
	$worksheet->write($a,21, $total_o_retail,$headNumber);
	$worksheet->write($a,22, $total_o_margin,$headNumber);
		if ($total_o_retail le 0){
			$worksheet->write($a,23, "",$headPct); }
		else{
			$worksheet->write($a,23, $total_o_margin/$total_o_retail,$headPct); }
			
	$worksheet->write($a,24, "",$headNumber);
	$worksheet->write($a,25, "",$headPct);
	$worksheet->write($a,26, "",$headNumber);
	$worksheet->write($a,27, "",$headPct);	
	
	$worksheet->write($a,28, $total_amt501000,$headNumber);
	$worksheet->write($a,29, "",$headNumber);
	$worksheet->write($a,30, "",$headNumber);
	
	$worksheet->write($a,31, $total_amt503200,$headNumber);
	$worksheet->write($a,32, "",$headNumber);
	$worksheet->write($a,33, "",$headNumber);
				
	$worksheet->write($a,34, $total_amt503250,$headNumber);
	$worksheet->write($a,35, "",$headNumber);
	$worksheet->write($a,36, "",$headNumber);
		
	$worksheet->write($a,37, $total_amt503500,$headNumber);
	$worksheet->write($a,38, "",$headNumber);
	$worksheet->write($a,39, "",$headNumber);
		
	$worksheet->write($a,40, $total_amt506000,$headNumber);
	$worksheet->write($a,41, "",$headNumber);
	$worksheet->write($a,42, "",$headNumber);
	
	$worksheet->write($a,43, $total_amt503000,$headNumber);
	$worksheet->write($a,44, "",$headNumber);
	$worksheet->write($a,45, "",$headNumber);
				
	$worksheet->write($a,46, $total_amt507000,$headNumber);
	$worksheet->write($a,47, "",$headNumber);
	$worksheet->write($a,48, "",$headNumber);
	
	$worksheet->write($a,49, $total_amt999998,$headNumber);
	$worksheet->write($a,50, "",$headNumber);
	$worksheet->write($a,51, "",$headNumber);
	
	$worksheet->write($a,52, $total_amt505000,$headNumber);
	$worksheet->write($a,53, "",$headNumber);
	$worksheet->write($a,54, "",$headNumber);
				
	$worksheet->write($a,55, $total_amt432000,$headNumber);
	$worksheet->write($a,56, "",$headNumber);
	$worksheet->write($a,57, "",$headNumber);
	
	$worksheet->write($a,58, $total_amt433000,$headNumber);
	$worksheet->write($a,59, "",$headNumber);
	$worksheet->write($a,60, "",$headNumber);
	
	$worksheet->write($a,61, $total_amt458490,$headNumber);
	$worksheet->write($a,62, "",$headNumber);
	$worksheet->write($a,63, "",$headNumber);
	
	$worksheet->write($a,64, $total_amt504000,$headNumber);
	$worksheet->write($a,65, "",$headNumber);
	$worksheet->write($a,66, "",$headNumber);
	
	$worksheet->write($a,67, $total_c_retail,$headNumber);
	$worksheet->write($a,68, $total_c_margin+$total_amt434000+$total_amt458550+$total_amt460100+$total_amt460200+$total_amt460300,$headNumber);
		if ($s->{c_retail} le 0){
			$worksheet->write($a,69, "",$headPct); }
		else{
			$worksheet->write($a,69, ($total_c_margin+$total_amt434000+$total_amt458550+$total_amt460100+$total_amt460200+$total_amt460300)/$total_c_retail,$headPct); }
			
	$worksheet->write($a,70, "",$headNumber);
	$worksheet->write($a,71, "",$headPct);
	$worksheet->write($a,72, "",$headNumber);
	$worksheet->write($a,73, "",$headPct);	
	
	$worksheet->write($a,74, $total_c_retail,$headNumber);
	$worksheet->write($a,75, $total_c_margin,$headNumber);
		if ($s->{c_retail} le 0){
			$worksheet->write($a,76, "",$headPct); }
		else{
			$worksheet->write($a,76, $total_c_margin/$total_c_retail,$headPct); }
		
	$worksheet->write($a,77, "",$headNumber);
	$worksheet->write($a,78, "",$headPct);
	$worksheet->write($a,79, "",$headNumber);
	$worksheet->write($a,80, "",$headPct);	
			
	$worksheet->write($a,81, $total_amt434000,$headNumber);
	$worksheet->write($a,82, "",$headNumber);
	$worksheet->write($a,83, "",$headNumber);
			
	$worksheet->write($a,84, $total_amt458550,$headNumber);
	$worksheet->write($a,85, "",$headNumber);
	$worksheet->write($a,86, "",$headNumber);
	
	$worksheet->write($a,87, $total_amt460100,$headNumber);
	$worksheet->write($a,88, "",$headNumber);
	$worksheet->write($a,89, "",$headNumber);
		
	$worksheet->write($a,90, $total_amt460200,$headNumber);
	$worksheet->write($a,91, "",$headNumber);
	$worksheet->write($a,92, "",$headNumber);
			
	$worksheet->write($a,93, $total_amt460300,$headNumber);
	$worksheet->write($a,94, "",$headNumber);
	$worksheet->write($a,95, "",$headNumber);		

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

$table = 'consolidated_margin_x.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls = $dbh_csv->prepare (qq{SELECT SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
								 FROM $table 
								 WHERE (merch_group_code_rev = 'DS' or merch_group_code_rev = 'SU') AND (((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
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
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40'))
								});
$sls->execute();


while(my $s = $sls->fetchrow_hashref()){
		$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
		$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
		$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
			if (($s->{o_retail}+$s->{c_retail}) <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $s->{o_retail},$border1); # outright sales retail
		$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin 
			if ($s->{o_retail} <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);			
		
		$worksheet->write($a,21, $s->{o_retail},$border1);
		$worksheet->write($a,22, $s->{o_margin},$border1);
			if ($s->{c_retail} <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
				
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);	
		
		$worksheet->write($a,28, $s->{amt501000},$border1);
		$worksheet->write($a,29, "",$border1);
		$worksheet->write($a,30, "",$border1);
		
		$worksheet->write($a,31, $s->{amt503200},$border1);
		$worksheet->write($a,32, "",$border1);
		$worksheet->write($a,33, "",$border1);
					
		$worksheet->write($a,34, $s->{amt503250},$border1);
		$worksheet->write($a,35, "",$border1);
		$worksheet->write($a,36, "",$border1);
			
		$worksheet->write($a,37, $s->{amt503500},$border1);
		$worksheet->write($a,38, "",$border1);
		$worksheet->write($a,39, "",$border1);
			
		$worksheet->write($a,40, $s->{amt506000},$border1);
		$worksheet->write($a,41, "",$border1);
		$worksheet->write($a,42, "",$border1);
		
		$worksheet->write($a,43, $s->{amt503000},$border1);
		$worksheet->write($a,44, "",$border1);
		$worksheet->write($a,45, "",$border1);
		
		$worksheet->write($a,46, $s->{amt507000},$border1);
		$worksheet->write($a,47, "",$border1);
		$worksheet->write($a,48, "",$border1);
		
		$worksheet->write($a,49, $s->{amt999998},$border1);
		$worksheet->write($a,50, "",$border1);
		$worksheet->write($a,51, "",$border1);
		
		$worksheet->write($a,52, $s->{amt505000},$border1);
		$worksheet->write($a,53, "",$border1);
		$worksheet->write($a,54, "",$border1);
					
		$worksheet->write($a,55, $s->{amt432000},$border1);
		$worksheet->write($a,56, "",$border1);
		$worksheet->write($a,57, "",$border1);
		
		$worksheet->write($a,58, $s->{amt433000},$border1);
		$worksheet->write($a,59, "",$border1);
		$worksheet->write($a,60, "",$border1);
		
		$worksheet->write($a,61, $s->{amt458490},$border1);
		$worksheet->write($a,62, "",$border1);
		$worksheet->write($a,63, "",$border1);
		
		$worksheet->write($a,64, $s->{amt504000},$border1);
		$worksheet->write($a,65, "",$border1);
		$worksheet->write($a,66, "",$border1);
		
		$worksheet->write($a,67, $s->{c_retail},$border1);
		$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,69, "",$subt); }
			else{
				$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
				
		$worksheet->write($a,70, "",$border1);
		$worksheet->write($a,71, "",$subt);
		$worksheet->write($a,72, "",$border1);
		$worksheet->write($a,73, "",$subt);	
		
		$worksheet->write($a,74, $s->{c_retail},$border1);
		$worksheet->write($a,75, $s->{c_margin},$border1);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,76, "",$subt); }
			else{
				$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
			
		$worksheet->write($a,77, "",$border1);
		$worksheet->write($a,78, "",$subt);
		$worksheet->write($a,79, "",$border1);
		$worksheet->write($a,80, "",$subt);	
				
		$worksheet->write($a,81, $s->{amt434000},$border1);
		$worksheet->write($a,82, "",$border1);
		$worksheet->write($a,83, "",$border1);
				
		$worksheet->write($a,84, $s->{amt458550},$border1);
		$worksheet->write($a,85, "",$border1);
		$worksheet->write($a,86, "",$border1);
		
		$worksheet->write($a,87, $s->{amt460100},$border1);
		$worksheet->write($a,88, "",$border1);
		$worksheet->write($a,89, "",$border1);
			
		$worksheet->write($a,90, $s->{amt460200},$border1);
		$worksheet->write($a,91, "",$border1);
		$worksheet->write($a,92, "",$border1);
				
		$worksheet->write($a,93, $s->{amt460300},$border1);
		$worksheet->write($a,94, "",$border1);
		$worksheet->write($a,95, "",$border1);	
				
	$a++;
	$counter++;
}

$sls->finish();

}

sub query_by_store {

$table = 'consolidated_margin_x.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;


$sls = $dbh_csv->prepare (qq{SELECT store_code, store_description, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
								 FROM $table 
								 WHERE (merch_group_code_rev = 'DS' or merch_group_code_rev = 'SU') AND (((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
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
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40'))
								 GROUP BY store_code, store_description 
								 ORDER BY store_code
								});
$sls->execute();

while(my $s = $sls->fetchrow_hashref()){
	
	if($s1f2_counter ne 2){
		$worksheet->write($a,5, $s->{store_code},$desc);
		$worksheet->write($a,6, $s->{store_description},$desc);
		$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1);
		$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
			if (($s->{o_retail}+$s->{c_retail}) <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $s->{o_retail},$border1);
		$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1);
			if ($s->{o_retail} <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);			
		
		$worksheet->write($a,21, $s->{o_retail},$border1);
		$worksheet->write($a,22, $s->{o_margin},$border1);
			if ($s->{c_retail} <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
				
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);	
		
		$worksheet->write($a,28, $s->{amt501000},$border1);
		$worksheet->write($a,29, "",$border1);
		$worksheet->write($a,30, "",$border1);
		
		$worksheet->write($a,31, $s->{amt503200},$border1);
		$worksheet->write($a,32, "",$border1);
		$worksheet->write($a,33, "",$border1);
					
		$worksheet->write($a,34, $s->{amt503250},$border1);
		$worksheet->write($a,35, "",$border1);
		$worksheet->write($a,36, "",$border1);
			
		$worksheet->write($a,37, $s->{amt503500},$border1);
		$worksheet->write($a,38, "",$border1);
		$worksheet->write($a,39, "",$border1);
			
		$worksheet->write($a,40, $s->{amt506000},$border1);
		$worksheet->write($a,41, "",$border1);
		$worksheet->write($a,42, "",$border1);
		
		$worksheet->write($a,43, $s->{amt503000},$border1);
		$worksheet->write($a,44, "",$border1);
		$worksheet->write($a,45, "",$border1);
					
		$worksheet->write($a,46, $s->{amt507000},$border1);
		$worksheet->write($a,47, "",$border1);
		$worksheet->write($a,48, "",$border1);
		
		$worksheet->write($a,49, $s->{amt999998},$border1);
		$worksheet->write($a,50, "",$border1);
		$worksheet->write($a,51, "",$border1);
		
		$worksheet->write($a,52, $s->{amt505000},$border1);
		$worksheet->write($a,53, "",$border1);
		$worksheet->write($a,54, "",$border1);
					
		$worksheet->write($a,55, $s->{amt432000},$border1);
		$worksheet->write($a,56, "",$border1);
		$worksheet->write($a,57, "",$border1);
		
		$worksheet->write($a,58, $s->{amt433000},$border1);
		$worksheet->write($a,59, "",$border1);
		$worksheet->write($a,60, "",$border1);
		
		$worksheet->write($a,61, $s->{amt458490},$border1);
		$worksheet->write($a,62, "",$border1);
		$worksheet->write($a,63, "",$border1);
		
		$worksheet->write($a,64, $s->{amt504000},$border1);
		$worksheet->write($a,65, "",$border1);
		$worksheet->write($a,66, "",$border1);
		
		$worksheet->write($a,67, $s->{c_retail},$border1);
		$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,69, "",$subt); }
			else{
				$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
				
		$worksheet->write($a,70, "",$border1);
		$worksheet->write($a,71, "",$subt);
		$worksheet->write($a,72, "",$border1);
		$worksheet->write($a,73, "",$subt);	
		
		$worksheet->write($a,74, $s->{c_retail},$border1);
		$worksheet->write($a,75, $s->{c_margin},$border1);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,76, "",$subt); }
			else{
				$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
			
		$worksheet->write($a,77, "",$border1);
		$worksheet->write($a,78, "",$subt);
		$worksheet->write($a,79, "",$border1);
		$worksheet->write($a,80, "",$subt);	
				
		$worksheet->write($a,81, $s->{amt434000},$border1);
		$worksheet->write($a,82, "",$border1);
		$worksheet->write($a,83, "",$border1);
				
		$worksheet->write($a,84, $s->{amt458550},$border1);
		$worksheet->write($a,85, "",$border1);
		$worksheet->write($a,86, "",$border1);
		
		$worksheet->write($a,87, $s->{amt460100},$border1);
		$worksheet->write($a,88, "",$border1);
		$worksheet->write($a,89, "",$border1);
			
		$worksheet->write($a,90, $s->{amt460200},$border1);
		$worksheet->write($a,91, "",$border1);
		$worksheet->write($a,92, "",$border1);
				
		$worksheet->write($a,93, $s->{amt460300},$border1);
		$worksheet->write($a,94, "",$border1);
		$worksheet->write($a,95, "",$border1);	
		
			if ($mrch1 eq 'SU' and $mrch2 eq 'SU' and ($s->{store_code} eq '2001' or $s->{store_code} eq '2001W')) {			
				$s1f2_o_retail += $s->{o_retail};
				$s1f2_o_margin += $s->{o_margin};
				$s1f2_c_retail += $s->{c_retail};
				$s1f2_c_margin += $s->{c_margin};
				$s1f2_amt501000 += $s->{amt501000};
				$s1f2_amt503200 += $s->{amt503200};
				$s1f2_amt503250 += $s->{amt503250};
				$s1f2_amt503500 += $s->{amt503500};
				$s1f2_amt506000 += $s->{amt506000};
				$s1f2_amt503000 += $s->{amt503000};
				$s1f2_amt507000 += $s->{amt507000};
				$s1f2_amt999998 += $s->{amt999998};
				$s1f2_amt504000 += $s->{amt504000};
				$s1f2_amt505000 += $s->{amt505000};
				$s1f2_amt432000 += $s->{amt432000};
				$s1f2_amt433000 += $s->{amt433000};
				$s1f2_amt458490 += $s->{amt458490};
				$s1f2_amt434000 += $s->{amt434000};
				$s1f2_amt458550 += $s->{amt458550};
				$s1f2_amt460100 += $s->{amt460100};
				$s1f2_amt460200 += $s->{amt460200};
				$s1f2_amt460300 += $s->{amt460300};
				
				$s1f2_counter ++; # once value = 2, we'll have a summation of s1 and f2
			} 
		
		$a++;
		$counter++;
		
	}
	
	if($s1f2_counter eq 2){
		$worksheet->write($a,5, "",$desc);
		$worksheet->write($a,6, "METRO COLON + F2",$desc);
		$worksheet->write($a,7, $s1f2_o_retail+$s1f2_c_retail,$border1);
		$worksheet->write($a,8, $s1f2_o_margin+$s1f2_c_margin+$s1f2_amt501000+$s1f2_amt432000+$s1f2_amt433000+$s1f2_amt458490+$s1f2_amt503200+$s1f2_amt503250+$s1f2_amt503500+$ss1f2_amt506000+$s1f2_amt503000+$s1f2_amt507000+$s1f2_amt999998+$s1f2_amt505000+$s1f2_amt504000+$s1f2_amt434000+$s1f2_amt458550+$s1f2_amt460100+$s1f2_amt460200+$s1f2_amt460300,$border1);
			if (($s1f2_o_retail+$s1f2_c_retail) <= 0){
				$worksheet->write($a,9, "",$subt); }
			else{
				$worksheet->write($a,9, ($s1f2_o_margin+$s1f2_c_margin+$s1f2_amt501000+$s1f2_amt432000+$s1f2_amt433000+$s1f2_amt458490+$s1f2_amt503200+$s1f2_amt503250+$s1f2_amt503500+$ss1f2_amt506000+$s1f2_amt503000+$s1f2_amt507000+$s1f2_amt999998+$s1f2_amt505000+$s1f2_amt504000+$s1f2_amt434000+$s1f2_amt458550+$s1f2_amt460100+$s1f2_amt460200+$s1f2_amt460300)/($s1f2_o_retail+$s1f2_c_retail),$subt); }
		
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$subt);
		$worksheet->write($a,12, "",$border1);
		$worksheet->write($a,13, "",$subt);
		
		$worksheet->write($a,14, $s1f2_o_retail,$border1);
		$worksheet->write($a,15, $s1f2_o_margin+$s1f2_amt501000+$s1f2_amt432000+$s1f2_amt433000+$s1f2_amt458490+$s1f2_amt503200+$s1f2_amt503250+$s1f2_amt503500+$ss1f2_amt506000+$s1f2_amt503000+$s1f2_amt507000+$s1f2_amt999998+$s1f2_amt505000+$s1f2_amt504000,$border1);
			if ($s1f2_o_retail <= 0){
				$worksheet->write($a,16, "",$subt); }
			else{
				$worksheet->write($a,16, ($s1f2_o_margin+$s1f2_amt501000+$s1f2_amt432000+$s1f2_amt433000+$s1f2_amt458490+$s1f2_amt503200+$s1f2_amt503250+$s1f2_amt503500+$ss1f2_amt506000+$s1f2_amt503000+$s1f2_amt507000+$s1f2_amt999998+$s1f2_amt505000+$s1f2_amt504000)/$s1f2_o_retail,$subt); }
		
		$worksheet->write($a,17, "",$border1);
		$worksheet->write($a,18, "",$subt);
		$worksheet->write($a,19, "",$border1);
		$worksheet->write($a,20, "",$subt);			
		
		$worksheet->write($a,21, $s1f2_o_retail,$border1);
		$worksheet->write($a,22, $s1f2_o_margin,$border1);
			if ($s1f2_c_retail <= 0){
				$worksheet->write($a,23, "",$subt); }
			else{
				$worksheet->write($a,23, $s1f2_o_margin/$s1f2_o_retail,$subt); }
				
		$worksheet->write($a,24, "",$border1);
		$worksheet->write($a,25, "",$subt);
		$worksheet->write($a,26, "",$border1);
		$worksheet->write($a,27, "",$subt);	
		
		$worksheet->write($a,28, $s1f2_amt501000,$border1);
		$worksheet->write($a,29, "",$border1);
		$worksheet->write($a,30, "",$border1);
		
		$worksheet->write($a,31, $s1f2_amt503200,$border1);
		$worksheet->write($a,32, "",$border1);
		$worksheet->write($a,33, "",$border1);
					
		$worksheet->write($a,34, $s1f2_amt503250,$border1);
		$worksheet->write($a,35, "",$border1);
		$worksheet->write($a,36, "",$border1);
			
		$worksheet->write($a,37, $s1f2_amt503500,$border1);
		$worksheet->write($a,38, "",$border1);
		$worksheet->write($a,39, "",$border1);
			
		$worksheet->write($a,40, $s1f2_amt506000,$border1);
		$worksheet->write($a,41, "",$border1);
		$worksheet->write($a,42, "",$border1);
		
		$worksheet->write($a,43, $s1f2_amt503000,$border1);
		$worksheet->write($a,44, "",$border1);
		$worksheet->write($a,45, "",$border1);
					
		$worksheet->write($a,46, $s1f2_amt507000,$border1);
		$worksheet->write($a,47, "",$border1);
		$worksheet->write($a,48, "",$border1);
		
		$worksheet->write($a,49, $s1f2_amt999998,$border1);
		$worksheet->write($a,50, "",$border1);
		$worksheet->write($a,51, "",$border1);
		
		$worksheet->write($a,52, $s1f2_amt505000,$border1);
		$worksheet->write($a,53, "",$border1);
		$worksheet->write($a,54, "",$border1);
					
		$worksheet->write($a,55, $s1f2_amt432000,$border1);
		$worksheet->write($a,56, "",$border1);
		$worksheet->write($a,57, "",$border1);
		
		$worksheet->write($a,58, $s1f2_amt433000,$border1);
		$worksheet->write($a,59, "",$border1);
		$worksheet->write($a,60, "",$border1);
		
		$worksheet->write($a,61, $s1f2_amt458490,$border1);
		$worksheet->write($a,62, "",$border1);
		$worksheet->write($a,63, "",$border1);
		
		$worksheet->write($a,64, $s1f2_amt504000,$border1);
		$worksheet->write($a,65, "",$border1);
		$worksheet->write($a,66, "",$border1);
		
		$worksheet->write($a,67, $s1f2_c_retail,$border1);
		$worksheet->write($a,68, $s1f2_c_margin+$s1f2_amt434000+$s1f2_amt458550+$s1f2_amt460100+$s1f2_amt460200+$s1f2_amt460300,$border1);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,69, "",$subt); }
			else{
				$worksheet->write($a,69, ($s1f2_c_margin+$s1f2_amt434000+$s1f2_amt458550+$s1f2_amt460100+$s1f2_amt460200+$s1f2_amt460300)/$s1f2_c_retail,$subt); }
				
		$worksheet->write($a,70, "",$border1);
		$worksheet->write($a,71, "",$subt);
		$worksheet->write($a,72, "",$border1);
		$worksheet->write($a,73, "",$subt);	
		
		$worksheet->write($a,74, $s1f2_c_retail,$border1);
		$worksheet->write($a,75, $s1f2_c_margin,$border1);
			if ($s->{c_retail} le 0){
				$worksheet->write($a,76, "",$subt); }
			else{
				$worksheet->write($a,76, $s1f2_c_margin/$s1f2_c_retail,$subt); }
			
		$worksheet->write($a,77, "",$border1);
		$worksheet->write($a,78, "",$subt);
		$worksheet->write($a,79, "",$border1);
		$worksheet->write($a,80, "",$subt);	
				
		$worksheet->write($a,81, $s1f2_amt434000,$border1);
		$worksheet->write($a,82, "",$border1);
		$worksheet->write($a,83, "",$border1);
				
		$worksheet->write($a,84, $s1f2_amt458550,$border1);
		$worksheet->write($a,85, "",$border1);
		$worksheet->write($a,86, "",$border1);
		
		$worksheet->write($a,87, $s1f2_amt460100,$border1);
		$worksheet->write($a,88, "",$border1);
		$worksheet->write($a,89, "",$border1);
			
		$worksheet->write($a,90, $s1f2_amt460200,$border1);
		$worksheet->write($a,91, "",$border1);
		$worksheet->write($a,92, "",$border1);
				
		$worksheet->write($a,93, $s1f2_amt460300,$border1);
		$worksheet->write($a,94, "",$border1);
		$worksheet->write($a,95, "",$border1);	
		
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

$table = 'consolidated_margin_x.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
	
$sls = $dbh_csv->prepare (qq{SELECT merch_group_code_rev, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
								 FROM $table 
								 WHERE (merch_group_code_rev = 'DS' or merch_group_code_rev = 'SU') AND (((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
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
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40'))
								GROUP BY merch_group_code_rev								
								});
$sls->execute();

if ($mrch1 eq 'DS' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){
	$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
	while(my $s = $sls->fetchrow_hashref()){
			if($s->{merch_group_code_rev} eq 'DS'){
				$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
				$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
					if (($s->{o_retail}+$s->{c_retail}) <= 0){
						$worksheet->write($a,9, "",$subt); }
					else{
						$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
				$worksheet->write($a,10, "",$border1);
				$worksheet->write($a,11, "",$subt);
				$worksheet->write($a,12, "",$border1);
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{o_retail},$border1); # outright sales retail
				$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
					if ($s->{o_retail} <= 0){
						$worksheet->write($a,16, "",$subt); }
					else{
						$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
				
				$worksheet->write($a,17, "",$border1);
				$worksheet->write($a,18, "",$subt);
				$worksheet->write($a,19, "",$border1);
				$worksheet->write($a,20, "",$subt);			
				
				$worksheet->write($a,21, $s->{o_retail},$border1);
				$worksheet->write($a,22, $s->{o_margin},$border1);
					if ($s->{c_retail} <= 0){
						$worksheet->write($a,23, "",$subt); }
					else{
						$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
						
				$worksheet->write($a,24, "",$border1);
				$worksheet->write($a,25, "",$subt);
				$worksheet->write($a,26, "",$border1);
				$worksheet->write($a,27, "",$subt);	
				
				$worksheet->write($a,28, $s->{amt501000},$border1);
				$worksheet->write($a,29, "",$border1);
				$worksheet->write($a,30, "",$border1);
				
				$worksheet->write($a,31, $s->{amt503200},$border1);
				$worksheet->write($a,32, "",$border1);
				$worksheet->write($a,33, "",$border1);
							
				$worksheet->write($a,34, $s->{amt503250},$border1);
				$worksheet->write($a,35, "",$border1);
				$worksheet->write($a,36, "",$border1);
					
				$worksheet->write($a,37, $s->{amt503500},$border1);
				$worksheet->write($a,38, "",$border1);
				$worksheet->write($a,39, "",$border1);
					
				$worksheet->write($a,40, "",$border1);
				$worksheet->write($a,41, "",$border1);
				$worksheet->write($a,42, "",$border1);
				
				$worksheet->write($a,43, $s->{amt506000},$border1);
				$worksheet->write($a,44, "",$border1);
				$worksheet->write($a,45, "",$border1);
							
				$worksheet->write($a,46, $s->{amt503000},$border1);
				$worksheet->write($a,47, "",$border1);
				$worksheet->write($a,48, "",$border1);
				
				$worksheet->write($a,49, $s->{amt507000},$border1);
				$worksheet->write($a,50, "",$border1);
				$worksheet->write($a,51, "",$border1);
				
				$worksheet->write($a,52, $s->{amt999998},$border1);
				$worksheet->write($a,53, "",$border1);
				$worksheet->write($a,54, "",$border1);
							
				$worksheet->write($a,55, $s->{amt432000},$border1);
				$worksheet->write($a,56, "",$border1);
				$worksheet->write($a,57, "",$border1);
				
				$worksheet->write($a,58, $s->{amt433000},$border1);
				$worksheet->write($a,59, "",$border1);
				$worksheet->write($a,60, "",$border1);
				
				$worksheet->write($a,61, $s->{amt458490},$border1);
				$worksheet->write($a,62, "",$border1);
				$worksheet->write($a,63, "",$border1);
				
				$worksheet->write($a,64, $s->{amt505000},$border1);
				$worksheet->write($a,65, "",$border1);
				$worksheet->write($a,66, "",$border1);
				
				$worksheet->write($a,67, $s->{c_retail},$border1);
				$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,69, "",$subt); }
					else{
						$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
						
				$worksheet->write($a,70, "",$border1);
				$worksheet->write($a,71, "",$subt);
				$worksheet->write($a,72, "",$border1);
				$worksheet->write($a,73, "",$subt);	
				
				$worksheet->write($a,74, $s->{c_retail},$border1);
				$worksheet->write($a,75, $s->{c_margin},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,76, "",$subt); }
					else{
						$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
					
				$worksheet->write($a,77, "",$border1);
				$worksheet->write($a,78, "",$subt);
				$worksheet->write($a,79, "",$border1);
				$worksheet->write($a,80, "",$subt);	
						
				$worksheet->write($a,81, $s->{amt434000},$border1);
				$worksheet->write($a,82, "",$border1);
				$worksheet->write($a,83, "",$border1);
						
				$worksheet->write($a,84, $s->{amt458550},$border1);
				$worksheet->write($a,85, "",$border1);
				$worksheet->write($a,86, "",$border1);
				
				$worksheet->write($a,87, $s->{amt460100},$border1);
				$worksheet->write($a,88, "",$border1);
				$worksheet->write($a,89, "",$border1);
					
				$worksheet->write($a,90, $s->{amt460200},$border1);
				$worksheet->write($a,91, "",$border1);
				$worksheet->write($a,92, "",$border1);
						
				$worksheet->write($a,93, $s->{amt460300},$border1);
				$worksheet->write($a,94, "",$border1);
				$worksheet->write($a,95, "",$border1);
			}
			
			else{	
				$a -= 1;
				$counter -= 1;
							
				$worksheet->write($a,96, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
				$worksheet->write($a,97, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
					if (($s->{o_retail}+$s->{c_retail}) <= 0){
						$worksheet->write($a,98, "",$subt); }
					else{
						$worksheet->write($a,98, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
					
				$worksheet->write($a,99, "",$border1);
				$worksheet->write($a,100, "",$subt);
				$worksheet->write($a,101, "",$border1);
				$worksheet->write($a,102, "",$subt);
				
				$worksheet->write($a,103, $s->{o_retail},$border1); # outright sales retail
				$worksheet->write($a,104, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
					if ($s->{o_retail} <= 0){
					$worksheet->write($a,105, "",$subt); }
					else{
					$worksheet->write($a,105, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
					
				$worksheet->write($a,106, "",$border1);
				$worksheet->write($a,107, "",$subt);
				$worksheet->write($a,108, "",$border1);
				$worksheet->write($a,109, "",$subt);			
				
				$worksheet->write($a,110, $s->{o_retail},$border1);
				$worksheet->write($a,111, $s->{o_margin},$border1);
					if ($s->{c_retail} <= 0){
						$worksheet->write($a,112, "",$subt); }
					else{
						$worksheet->write($a,112, $s->{o_margin}/$s->{o_retail},$subt); }
						
				$worksheet->write($a,113, "",$border1);
				$worksheet->write($a,114, "",$subt);
				$worksheet->write($a,115, "",$border1);
				$worksheet->write($a,116, "",$subt);	
				
				$worksheet->write($a,117, $s->{amt501000},$border1);
				$worksheet->write($a,118, "",$border1);
				$worksheet->write($a,119, "",$border1);
				
				$worksheet->write($a,120, $s->{amt503200},$border1);
				$worksheet->write($a,121, "",$border1);
				$worksheet->write($a,122, "",$border1);
							
				$worksheet->write($a,123, $s->{amt503250},$border1);
				$worksheet->write($a,124, "",$border1);
				$worksheet->write($a,125, "",$border1);
					
				$worksheet->write($a,126, $s->{amt503500},$border1);
				$worksheet->write($a,127, "",$border1);
				$worksheet->write($a,128, "",$border1);
					
				$worksheet->write($a,129, "",$border1);
				$worksheet->write($a,130, "",$border1);
				$worksheet->write($a,131, "",$border1);
				
				$worksheet->write($a,132, $s->{amt506000},$border1);
				$worksheet->write($a,133, "",$border1);
				$worksheet->write($a,134, "",$border1);
						
				$worksheet->write($a,135, $s->{amt503000},$border1);
				$worksheet->write($a,136, "",$border1);
				$worksheet->write($a,137, "",$border1);
				
				$worksheet->write($a,138, $s->{amt507000},$border1);
				$worksheet->write($a,139, "",$border1);
				$worksheet->write($a,140, "",$border1);
				
				$worksheet->write($a,141, $s->{amt999998},$border1);
				$worksheet->write($a,142, "",$border1);
				$worksheet->write($a,143, "",$border1);
							
				$worksheet->write($a,144, $s->{amt432000},$border1);
				$worksheet->write($a,145, "",$border1);
				$worksheet->write($a,146, "",$border1);
			
				$worksheet->write($a,147, $s->{amt433000},$border1);
				$worksheet->write($a,148, "",$border1);
				$worksheet->write($a,149, "",$border1);
				
				$worksheet->write($a,150, $s->{amt458490},$border1);
				$worksheet->write($a,151, "",$border1);
				$worksheet->write($a,152, "",$border1);
				
				$worksheet->write($a,153, $s->{amt505000},$border1);
				$worksheet->write($a,154, "",$border1);
				$worksheet->write($a,155, "",$border1);
				
				$worksheet->write($a,156, $s->{c_retail},$border1);
				$worksheet->write($a,157, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,158, "",$subt); }
					else{
						$worksheet->write($a,158, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
						
				$worksheet->write($a,159, "",$border1);
				$worksheet->write($a,160, "",$subt);
				$worksheet->write($a,161, "",$border1);
				$worksheet->write($a,162, "",$subt);	
				
				$worksheet->write($a,163, $s->{c_retail},$border1);
				$worksheet->write($a,164, $s->{c_margin},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,165, "",$subt); }
					else{
						$worksheet->write($a,165, $s->{c_margin}/$s->{c_retail},$subt); }
					
				$worksheet->write($a,166, "",$border1);
				$worksheet->write($a,167, "",$subt);
				$worksheet->write($a,168, "",$border1);
				$worksheet->write($a,169, "",$subt);	
						
				$worksheet->write($a,170, $s->{amt434000},$border1);
				$worksheet->write($a,171, "",$border1);
				$worksheet->write($a,172, "",$border1);
						
				$worksheet->write($a,173, $s->{amt458550},$border1);
				$worksheet->write($a,174, "",$border1);
				$worksheet->write($a,175, "",$border1);
				
				$worksheet->write($a,176, $s->{amt460100},$border1);
				$worksheet->write($a,177, "",$border1);
				$worksheet->write($a,178, "",$border1);
					
				$worksheet->write($a,179, $s->{amt460200},$border1);
				$worksheet->write($a,180, "",$border1);
				$worksheet->write($a,181, "",$border1);
						
				$worksheet->write($a,182, $s->{amt460300},$border1);
				$worksheet->write($a,183, "",$border1);
				$worksheet->write($a,184, "",$border1);
			}
			
			$a++;
			$counter++;			
	}
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'SU' and $mrch3 eq 'OT' ){
	$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
	while(my $s = $sls->fetchrow_hashref()){
		$worksheet->write($a,5, $s->{store_code},$desc);
		$worksheet->write($a,6, $s->{store_description},$desc);
		
		if($s->{merch_group_code_rev} eq 'DS'){ 		
			$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
			$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
				if (($s->{o_retail}+$s->{c_retail}) <= 0){
					$worksheet->write($a,9, "",$subt); }
				else{
					$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
			$worksheet->write($a,10, "",$border1);
			$worksheet->write($a,11, "",$subt);
			$worksheet->write($a,12, "",$border1);
			$worksheet->write($a,13, "",$subt);
				
			$worksheet->write($a,14, $s->{o_retail},$border1); # outright sales retail
			$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
				if ($s->{o_retail} <= 0){
					$worksheet->write($a,16, "",$subt); }
				else{
					$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
				
			$worksheet->write($a,17, "",$border1);
			$worksheet->write($a,18, "",$subt);
			$worksheet->write($a,19, "",$border1);
			$worksheet->write($a,20, "",$subt);			
			
			$worksheet->write($a,21, $s->{o_retail},$border1);
			$worksheet->write($a,22, $s->{o_margin},$border1);
				if ($s->{c_retail} <= 0){
					$worksheet->write($a,23, "",$subt); }
				else{
					$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
					
			$worksheet->write($a,24, "",$border1);
			$worksheet->write($a,25, "",$subt);
			$worksheet->write($a,26, "",$border1);
			$worksheet->write($a,27, "",$subt);	
			
			$worksheet->write($a,28, $s->{amt501000},$border1);
			$worksheet->write($a,29, "",$border1);
			$worksheet->write($a,30, "",$border1);
			
			$worksheet->write($a,31, $s->{amt503200},$border1);
			$worksheet->write($a,32, "",$border1);
			$worksheet->write($a,33, "",$border1);
						
			$worksheet->write($a,34, $s->{amt503250},$border1);
			$worksheet->write($a,35, "",$border1);
			$worksheet->write($a,36, "",$border1);
				
			$worksheet->write($a,37, $s->{amt503500},$border1);
			$worksheet->write($a,38, "",$border1);
			$worksheet->write($a,39, "",$border1);
				
			$worksheet->write($a,40, "",$border1);
			$worksheet->write($a,41, "",$border1);
			$worksheet->write($a,42, "",$border1);
			
			$worksheet->write($a,43, $s->{amt506000},$border1);
			$worksheet->write($a,44, "",$border1);
			$worksheet->write($a,45, "",$border1);
						
			$worksheet->write($a,46, $s->{amt503000},$border1);
			$worksheet->write($a,47, "",$border1);
			$worksheet->write($a,48, "",$border1);
			
			$worksheet->write($a,49, $s->{amt507000},$border1);
			$worksheet->write($a,50, "",$border1);
			$worksheet->write($a,51, "",$border1);
			
			$worksheet->write($a,52, $s->{amt999998},$border1);
			$worksheet->write($a,53, "",$border1);
			$worksheet->write($a,54, "",$border1);
						
			$worksheet->write($a,55, $s->{amt432000},$border1);
			$worksheet->write($a,56, "",$border1);
			$worksheet->write($a,57, "",$border1);
			
			$worksheet->write($a,58, $s->{amt433000},$border1);
			$worksheet->write($a,59, "",$border1);
			$worksheet->write($a,60, "",$border1);
			
			$worksheet->write($a,61, $s->{amt458490},$border1);
			$worksheet->write($a,62, "",$border1);
			$worksheet->write($a,63, "",$border1);
			
			$worksheet->write($a,64, $s->{amt505000},$border1);
			$worksheet->write($a,65, "",$border1);
			$worksheet->write($a,66, "",$border1);
			
			$worksheet->write($a,67, $s->{c_retail},$border1);
			$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,69, "",$subt); }
				else{
					$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
					
			$worksheet->write($a,70, "",$border1);
			$worksheet->write($a,71, "",$subt);
			$worksheet->write($a,72, "",$border1);
			$worksheet->write($a,73, "",$subt);	
			
			$worksheet->write($a,74, $s->{c_retail},$border1);
			$worksheet->write($a,75, $s->{c_margin},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,76, "",$subt); }
				else{
					$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
				
			$worksheet->write($a,77, "",$border1);
			$worksheet->write($a,78, "",$subt);
			$worksheet->write($a,79, "",$border1);
			$worksheet->write($a,80, "",$subt);	
					
			$worksheet->write($a,81, $s->{amt434000},$border1);
			$worksheet->write($a,82, "",$border1);
			$worksheet->write($a,83, "",$border1);
					
			$worksheet->write($a,84, $s->{amt458550},$border1);
			$worksheet->write($a,85, "",$border1);
			$worksheet->write($a,86, "",$border1);
			
			$worksheet->write($a,87, $s->{amt460100},$border1);
			$worksheet->write($a,88, "",$border1);
			$worksheet->write($a,89, "",$border1);
				
			$worksheet->write($a,90, $s->{amt460200},$border1);
			$worksheet->write($a,91, "",$border1);
			$worksheet->write($a,92, "",$border1);
					
			$worksheet->write($a,93, $s->{amt460300},$border1);
			$worksheet->write($a,94, "",$border1);
			$worksheet->write($a,95, "",$border1);

			$a -= 1;
			$counter -= 1;	
			}
			
		else{						
			$worksheet->write($a,96, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
			$worksheet->write($a,97, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
				if (($s->{o_retail}+$s->{c_retail}) <= 0){
					$worksheet->write($a,98, "",$subt); }
				else{
					$worksheet->write($a,98, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
					
			$worksheet->write($a,99, "",$border1);
			$worksheet->write($a,100, "",$subt);
			$worksheet->write($a,101, "",$border1);
			$worksheet->write($a,102, "",$subt);
			
			$worksheet->write($a,103, $s->{o_retail},$border1); # outright sales retail
			$worksheet->write($a,104, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
				if ($s->{o_retail} <= 0){
					$worksheet->write($a,105, "",$subt); }
				else{
					$worksheet->write($a,105, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
					
			$worksheet->write($a,106, "",$border1);
			$worksheet->write($a,107, "",$subt);
			$worksheet->write($a,108, "",$border1);
			$worksheet->write($a,109, "",$subt);			
			
			$worksheet->write($a,110, $s->{o_retail},$border1);
			$worksheet->write($a,111, $s->{o_margin},$border1);
				if ($s->{c_retail} <= 0){
					$worksheet->write($a,112, "",$subt); }
				else{
					$worksheet->write($a,112, $s->{o_margin}/$s->{o_retail},$subt); }
					
			$worksheet->write($a,113, "",$border1);
			$worksheet->write($a,114, "",$subt);
			$worksheet->write($a,115, "",$border1);
			$worksheet->write($a,116, "",$subt);	
			
			$worksheet->write($a,117, $s->{amt501000},$border1);
			$worksheet->write($a,118, "",$border1);
			$worksheet->write($a,119, "",$border1);
			
			$worksheet->write($a,120, $s->{amt503200},$border1);
			$worksheet->write($a,121, "",$border1);
			$worksheet->write($a,122, "",$border1);
						
			$worksheet->write($a,123, $s->{amt503250},$border1);
			$worksheet->write($a,124, "",$border1);
			$worksheet->write($a,125, "",$border1);
				
			$worksheet->write($a,126, $s->{amt503500},$border1);
			$worksheet->write($a,127, "",$border1);
			$worksheet->write($a,128, "",$border1);
				
			$worksheet->write($a,129, "",$border1);
			$worksheet->write($a,130, "",$border1);
			$worksheet->write($a,131, "",$border1);
			
			$worksheet->write($a,132, $s->{amt506000},$border1);
			$worksheet->write($a,133, "",$border1);
			$worksheet->write($a,134, "",$border1);
					
			$worksheet->write($a,135, $s->{amt503000},$border1);
			$worksheet->write($a,136, "",$border1);
			$worksheet->write($a,137, "",$border1);
			
			$worksheet->write($a,138, $s->{amt507000},$border1);
			$worksheet->write($a,139, "",$border1);
			$worksheet->write($a,140, "",$border1);
			
			$worksheet->write($a,141, $s->{amt999998},$border1);
			$worksheet->write($a,142, "",$border1);
			$worksheet->write($a,143, "",$border1);
						
			$worksheet->write($a,144, $s->{amt432000},$border1);
			$worksheet->write($a,145, "",$border1);
			$worksheet->write($a,146, "",$border1);
			
			$worksheet->write($a,147, $s->{amt433000},$border1);
			$worksheet->write($a,148, "",$border1);
			$worksheet->write($a,149, "",$border1);
			
			$worksheet->write($a,150, $s->{amt458490},$border1);
			$worksheet->write($a,151, "",$border1);
			$worksheet->write($a,152, "",$border1);
			
			$worksheet->write($a,153, $s->{amt505000},$border1);
			$worksheet->write($a,154, "",$border1);
			$worksheet->write($a,155, "",$border1);
			
			$worksheet->write($a,156, $s->{c_retail},$border1);
			$worksheet->write($a,157, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,158, "",$subt); }
				else{
					$worksheet->write($a,158, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
						
			$worksheet->write($a,159, "",$border1);
			$worksheet->write($a,160, "",$subt);
			$worksheet->write($a,161, "",$border1);
			$worksheet->write($a,162, "",$subt);	
			
			$worksheet->write($a,163, $s->{c_retail},$border1);
			$worksheet->write($a,164, $s->{c_margin},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,165, "",$subt); }
				else{
					$worksheet->write($a,165, $s->{c_margin}/$s->{c_retail},$subt); }
				
			$worksheet->write($a,166, "",$border1);
			$worksheet->write($a,167, "",$subt);
			$worksheet->write($a,168, "",$border1);
			$worksheet->write($a,169, "",$subt);	
					
			$worksheet->write($a,170, $s->{amt434000},$border1);
			$worksheet->write($a,171, "",$border1);
			$worksheet->write($a,172, "",$border1);
					
			$worksheet->write($a,173, $s->{amt458550},$border1);
			$worksheet->write($a,174, "",$border1);
			$worksheet->write($a,175, "",$border1);
			
			$worksheet->write($a,176, $s->{amt460100},$border1);
			$worksheet->write($a,177, "",$border1);
			$worksheet->write($a,178, "",$border1);
					
			$worksheet->write($a,179, $s->{amt460200},$border1);
			$worksheet->write($a,180, "",$border1);
			$worksheet->write($a,181, "",$border1);
					
			$worksheet->write($a,182, $s->{amt460300},$border1);
			$worksheet->write($a,183, "",$border1);
			$worksheet->write($a,184, "",$border1);
			}		
			
		$a++;
		$counter++;		
	}
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){
	$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
	
	foreach my $i( 7..95 ){
		$worksheet->write($a,$i, "",$border1); }
	
	$sls_2 = $dbh_csv->prepare (qq{SELECT merch_group_code_rev, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
								 FROM $table 
								 WHERE (merch_group_code_rev = 'DS' or merch_group_code_rev = 'SU') AND (((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
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
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40'))
								GROUP BY merch_group_code_rev
								ORDER BY merch_group_code_rev
								});
	$sls_2->execute();

	while(my $s = $sls_2->fetchrow_hashref()){
	
		if($s->{merch_group_code_rev} eq 'DS'){
			$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
			$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
				if (($s->{o_retail}+$s->{c_retail}) <= 0){
					$worksheet->write($a,9, "",$subt); }
				else{
					$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
			
			$worksheet->write($a,10, "",$border1);
			$worksheet->write($a,11, "",$subt);
			$worksheet->write($a,12, "",$border1);
			$worksheet->write($a,13, "",$subt);
			
			$worksheet->write($a,14, $s->{o_retail},$border1); # outright sales retail
			$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
				if ($s->{o_retail} <= 0){
					$worksheet->write($a,16, "",$subt); }
				else{
					$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
			
			$worksheet->write($a,17, "",$border1);
			$worksheet->write($a,18, "",$subt);
			$worksheet->write($a,19, "",$border1);
			$worksheet->write($a,20, "",$subt);			
			
			$worksheet->write($a,21, $s->{o_retail},$border1);
			$worksheet->write($a,22, $s->{o_margin},$border1);
				if ($s->{c_retail} <= 0){
					$worksheet->write($a,23, "",$subt); }
				else{
					$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
					
			$worksheet->write($a,24, "",$border1);
			$worksheet->write($a,25, "",$subt);
			$worksheet->write($a,26, "",$border1);
			$worksheet->write($a,27, "",$subt);	
			
			$worksheet->write($a,28, $s->{amt501000},$border1);
			$worksheet->write($a,29, "",$border1);
			$worksheet->write($a,30, "",$border1);
			
			$worksheet->write($a,31, $s->{amt503200},$border1);
			$worksheet->write($a,32, "",$border1);
			$worksheet->write($a,33, "",$border1);
						
			$worksheet->write($a,34, $s->{amt503250},$border1);
			$worksheet->write($a,35, "",$border1);
			$worksheet->write($a,36, "",$border1);
				
			$worksheet->write($a,37, $s->{amt503500},$border1);
			$worksheet->write($a,38, "",$border1);
			$worksheet->write($a,39, "",$border1);
				
			$worksheet->write($a,40, $s->{amt506000},$border1);
			$worksheet->write($a,41, "",$border1);
			$worksheet->write($a,42, "",$border1);
			
			$worksheet->write($a,43, $s->{amt503000},$border1);
			$worksheet->write($a,44, "",$border1);
			$worksheet->write($a,45, "",$border1);
						
			$worksheet->write($a,46, $s->{amt507000},$border1);
			$worksheet->write($a,47, "",$border1);
			$worksheet->write($a,48, "",$border1);
			
			$worksheet->write($a,49, $s->{amt999998},$border1);
			$worksheet->write($a,50, "",$border1);
			$worksheet->write($a,51, "",$border1);
			
			$worksheet->write($a,52, $s->{amt505000},$border1);
			$worksheet->write($a,53, "",$border1);
			$worksheet->write($a,54, "",$border1);
						
			$worksheet->write($a,55, $s->{amt432000},$border1);
			$worksheet->write($a,56, "",$border1);
			$worksheet->write($a,57, "",$border1);
			
			$worksheet->write($a,58, $s->{amt433000},$border1);
			$worksheet->write($a,59, "",$border1);
			$worksheet->write($a,60, "",$border1);
			
			$worksheet->write($a,61, $s->{amt458490},$border1);
			$worksheet->write($a,62, "",$border1);
			$worksheet->write($a,63, "",$border1);
			
			$worksheet->write($a,64, $s->{amt504000},$border1);
			$worksheet->write($a,65, "",$border1);
			$worksheet->write($a,66, "",$border1);
			
			$worksheet->write($a,67, $s->{c_retail},$border1);
			$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,69, "",$subt); }
				else{
					$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
					
			$worksheet->write($a,70, "",$border1);
			$worksheet->write($a,71, "",$subt);
			$worksheet->write($a,72, "",$border1);
			$worksheet->write($a,73, "",$subt);	
			
			$worksheet->write($a,74, $s->{c_retail},$border1);
			$worksheet->write($a,75, $s->{c_margin},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,76, "",$subt); }
				else{
					$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
				
			$worksheet->write($a,77, "",$border1);
			$worksheet->write($a,78, "",$subt);
			$worksheet->write($a,79, "",$border1);
			$worksheet->write($a,80, "",$subt);	
					
			$worksheet->write($a,81, $s->{amt434000},$border1);
			$worksheet->write($a,82, "",$border1);
			$worksheet->write($a,83, "",$border1);
					
			$worksheet->write($a,84, $s->{amt458550},$border1);
			$worksheet->write($a,85, "",$border1);
			$worksheet->write($a,86, "",$border1);
			
			$worksheet->write($a,87, $s->{amt460100},$border1);
			$worksheet->write($a,88, "",$border1);
			$worksheet->write($a,89, "",$border1);
				
			$worksheet->write($a,90, $s->{amt460200},$border1);
			$worksheet->write($a,91, "",$border1);
			$worksheet->write($a,92, "",$border1);
					
			$worksheet->write($a,93, $s->{amt460300},$border1);
			$worksheet->write($a,94, "",$border1);
			$worksheet->write($a,95, "",$border1);
		}
		
		else{		
			$worksheet->write($a,96, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
			$worksheet->write($a,97, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
				if (($s->{o_retail}+$s->{c_retail}) <= 0){
					$worksheet->write($a,98, "",$subt); }
				else{
					$worksheet->write($a,98, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
			$worksheet->write($a,99, "",$border1);
			$worksheet->write($a,100, "",$subt);
			$worksheet->write($a,101, "",$border1);
			$worksheet->write($a,102, "",$subt);
			
			$worksheet->write($a,103, $s->{o_retail},$border1); # outright sales retail
			$worksheet->write($a,104, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
				if ($s->{o_retail} <= 0){
				$worksheet->write($a,105, "",$subt); }
				else{
				$worksheet->write($a,105, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
				
			$worksheet->write($a,106, "",$border1);
			$worksheet->write($a,107, "",$subt);
			$worksheet->write($a,108, "",$border1);
			$worksheet->write($a,109, "",$subt);			
			
			$worksheet->write($a,110, $s->{o_retail},$border1);
			$worksheet->write($a,111, $s->{o_margin},$border1);
				if ($s->{c_retail} <= 0){
					$worksheet->write($a,112, "",$subt); }
				else{
					$worksheet->write($a,112, $s->{o_margin}/$s->{o_retail},$subt); }
					
			$worksheet->write($a,113, "",$border1);
			$worksheet->write($a,114, "",$subt);
			$worksheet->write($a,115, "",$border1);
			$worksheet->write($a,116, "",$subt);	
			
			$worksheet->write($a,117, $s->{amt501000},$border1);
			$worksheet->write($a,118, "",$border1);
			$worksheet->write($a,119, "",$border1);
			
			$worksheet->write($a,120, $s->{amt503200},$border1);
			$worksheet->write($a,121, "",$border1);
			$worksheet->write($a,122, "",$border1);
						
			$worksheet->write($a,123, $s->{amt503250},$border1);
			$worksheet->write($a,124, "",$border1);
			$worksheet->write($a,125, "",$border1);
				
			$worksheet->write($a,126, $s->{amt503500},$border1);
			$worksheet->write($a,127, "",$border1);
			$worksheet->write($a,128, "",$border1);
				
			$worksheet->write($a,129, $s->{amt506000},$border1);
			$worksheet->write($a,130, "",$border1);
			$worksheet->write($a,131, "",$border1);
			
			$worksheet->write($a,132, $s->{amt503000},$border1);
			$worksheet->write($a,133, "",$border1);
			$worksheet->write($a,134, "",$border1);
					
			$worksheet->write($a,135, $s->{amt507000},$border1);
			$worksheet->write($a,136, "",$border1);
			$worksheet->write($a,137, "",$border1);
			
			$worksheet->write($a,138, $s->{amt999998},$border1);
			$worksheet->write($a,139, "",$border1);
			$worksheet->write($a,140, "",$border1);
			
			$worksheet->write($a,141, $s->{amt505000},$border1);
			$worksheet->write($a,142, "",$border1);
			$worksheet->write($a,143, "",$border1);
						
			$worksheet->write($a,144, $s->{amt432000},$border1);
			$worksheet->write($a,145, "",$border1);
			$worksheet->write($a,146, "",$border1);
		
			$worksheet->write($a,147, $s->{amt433000},$border1);
			$worksheet->write($a,148, "",$border1);
			$worksheet->write($a,149, "",$border1);
			
			$worksheet->write($a,150, $s->{amt458490},$border1);
			$worksheet->write($a,151, "",$border1);
			$worksheet->write($a,152, "",$border1);
			
			$worksheet->write($a,153, $s->{amt504000},$border1);
			$worksheet->write($a,154, "",$border1);
			$worksheet->write($a,155, "",$border1);
			
			$worksheet->write($a,156, $s->{c_retail},$border1);
			$worksheet->write($a,157, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,158, "",$subt); }
				else{
					$worksheet->write($a,158, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
					
			$worksheet->write($a,159, "",$border1);
			$worksheet->write($a,160, "",$subt);
			$worksheet->write($a,161, "",$border1);
			$worksheet->write($a,162, "",$subt);	
			
			$worksheet->write($a,163, $s->{c_retail},$border1);
			$worksheet->write($a,164, $s->{c_margin},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,165, "",$subt); }
				else{
					$worksheet->write($a,165, $s->{c_margin}/$s->{c_retail},$subt); }
				
			$worksheet->write($a,166, "",$border1);
			$worksheet->write($a,167, "",$subt);
			$worksheet->write($a,168, "",$border1);
			$worksheet->write($a,169, "",$subt);	
					
			$worksheet->write($a,170, $s->{amt434000},$border1);
			$worksheet->write($a,171, "",$border1);
			$worksheet->write($a,172, "",$border1);
					
			$worksheet->write($a,173, $s->{amt458550},$border1);
			$worksheet->write($a,174, "",$border1);
			$worksheet->write($a,175, "",$border1);
			
			$worksheet->write($a,176, $s->{amt460100},$border1);
			$worksheet->write($a,177, "",$border1);
			$worksheet->write($a,178, "",$border1);
				
			$worksheet->write($a,179, $s->{amt460200},$border1);
			$worksheet->write($a,180, "",$border1);
			$worksheet->write($a,181, "",$border1);
					
			$worksheet->write($a,182, $s->{amt460300},$border1);
			$worksheet->write($a,183, "",$border1);
			$worksheet->write($a,184, "",$border1);
		}
		
	}
	
	$sls_2->finish();
	
	$a++;
	$counter++;
}

$sls->finish();

}

sub query_by_store_merchandise {

$table = 'consolidated_margin_x.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;

$blank = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, hidden => 1 );
$worksheet->conditional_formatting( 'H32:FE60', { type     => 'cell',  criteria => '=', value    => 0, format   => $blank });	
#$worksheet->conditional_formatting( 'F9:AK2000',  { type => 'cell', criteria => '<', value => 0, format => $down } );			

$sls = $dbh_csv->prepare (qq{SELECT store_code, store_description, merch_group_code_rev, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
								 FROM $table 
								 WHERE (merch_group_code_rev = 'DS' or merch_group_code_rev = 'SU') AND (((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
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
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40'))
								 GROUP BY store_code, store_description, merch_group_code_rev
								 ORDER BY store_code, merch_group_code_rev
								});
$sls->execute();

if ($mrch1 eq 'DS' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){	
	while(my $s = $sls->fetchrow_hashref()){
			$worksheet->write($a,5, $s->{store_code},$desc);
			$worksheet->write($a,6, $s->{store_description},$desc);
			
			if($s->{merch_group_code_rev} eq 'DS'){
				$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
				$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
					if (($s->{o_retail}+$s->{c_retail}) <= 0){
						$worksheet->write($a,9, "",$subt); }
					else{
						$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
				$worksheet->write($a,10, "",$border1);
				$worksheet->write($a,11, "",$subt);
				$worksheet->write($a,12, "",$border1);
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{o_retail},$border1); # outright sales retail
				$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
					if ($s->{o_retail} <= 0){
						$worksheet->write($a,16, "",$subt); }
					else{
						$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
				
				$worksheet->write($a,17, "",$border1);
				$worksheet->write($a,18, "",$subt);
				$worksheet->write($a,19, "",$border1);
				$worksheet->write($a,20, "",$subt);			
				
				$worksheet->write($a,21, $s->{o_retail},$border1);
				$worksheet->write($a,22, $s->{o_margin},$border1);
					if ($s->{c_retail} <= 0){
						$worksheet->write($a,23, "",$subt); }
					else{
						$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
						
				$worksheet->write($a,24, "",$border1);
				$worksheet->write($a,25, "",$subt);
				$worksheet->write($a,26, "",$border1);
				$worksheet->write($a,27, "",$subt);	
				
				$worksheet->write($a,28, $s->{amt501000},$border1);
				$worksheet->write($a,29, "",$border1);
				$worksheet->write($a,30, "",$border1);
				
				$worksheet->write($a,31, $s->{amt503200},$border1);
				$worksheet->write($a,32, "",$border1);
				$worksheet->write($a,33, "",$border1);
							
				$worksheet->write($a,34, $s->{amt503250},$border1);
				$worksheet->write($a,35, "",$border1);
				$worksheet->write($a,36, "",$border1);
					
				$worksheet->write($a,37, $s->{amt503500},$border1);
				$worksheet->write($a,38, "",$border1);
				$worksheet->write($a,39, "",$border1);
					
				$worksheet->write($a,40, $s->{amt506000},$border1);
				$worksheet->write($a,41, "",$border1);
				$worksheet->write($a,42, "",$border1);
				
				$worksheet->write($a,43, $s->{amt503000},$border1);
				$worksheet->write($a,44, "",$border1);
				$worksheet->write($a,45, "",$border1);
							
				$worksheet->write($a,46, $s->{amt507000},$border1);
				$worksheet->write($a,47, "",$border1);
				$worksheet->write($a,48, "",$border1);
				
				$worksheet->write($a,49, $s->{amt999998},$border1);
				$worksheet->write($a,50, "",$border1);
				$worksheet->write($a,51, "",$border1);
				
				$worksheet->write($a,52, $s->{amt505000},$border1);
				$worksheet->write($a,53, "",$border1);
				$worksheet->write($a,54, "",$border1);
							
				$worksheet->write($a,55, $s->{amt432000},$border1);
				$worksheet->write($a,56, "",$border1);
				$worksheet->write($a,57, "",$border1);
				
				$worksheet->write($a,58, $s->{amt433000},$border1);
				$worksheet->write($a,59, "",$border1);
				$worksheet->write($a,60, "",$border1);
				
				$worksheet->write($a,61, $s->{amt458490},$border1);
				$worksheet->write($a,62, "",$border1);
				$worksheet->write($a,63, "",$border1);
				
				$worksheet->write($a,64, $s->{amt504000},$border1);
				$worksheet->write($a,65, "",$border1);
				$worksheet->write($a,66, "",$border1);
				
				$worksheet->write($a,67, $s->{c_retail},$border1);
				$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,69, "",$subt); }
					else{
						$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
						
				$worksheet->write($a,70, "",$border1);
				$worksheet->write($a,71, "",$subt);
				$worksheet->write($a,72, "",$border1);
				$worksheet->write($a,73, "",$subt);	
				
				$worksheet->write($a,74, $s->{c_retail},$border1);
				$worksheet->write($a,75, $s->{c_margin},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,76, "",$subt); }
					else{
						$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
					
				$worksheet->write($a,77, "",$border1);
				$worksheet->write($a,78, "",$subt);
				$worksheet->write($a,79, "",$border1);
				$worksheet->write($a,80, "",$subt);	
						
				$worksheet->write($a,81, $s->{amt434000},$border1);
				$worksheet->write($a,82, "",$border1);
				$worksheet->write($a,83, "",$border1);
						
				$worksheet->write($a,84, $s->{amt458550},$border1);
				$worksheet->write($a,85, "",$border1);
				$worksheet->write($a,86, "",$border1);
				
				$worksheet->write($a,87, $s->{amt460100},$border1);
				$worksheet->write($a,88, "",$border1);
				$worksheet->write($a,89, "",$border1);
					
				$worksheet->write($a,90, $s->{amt460200},$border1);
				$worksheet->write($a,91, "",$border1);
				$worksheet->write($a,92, "",$border1);
						
				$worksheet->write($a,93, $s->{amt460300},$border1);
				$worksheet->write($a,94, "",$border1);
				$worksheet->write($a,95, "",$border1);
			}
			
			else{	
				$a -= 1;
				$counter -= 1;
							
				$worksheet->write($a,96, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
				$worksheet->write($a,97, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
					if (($s->{o_retail}+$s->{c_retail}) <= 0){
						$worksheet->write($a,98, "",$subt); }
					else{
						$worksheet->write($a,98, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
					
				$worksheet->write($a,99, "",$border1);
				$worksheet->write($a,100, "",$subt);
				$worksheet->write($a,101, "",$border1);
				$worksheet->write($a,102, "",$subt);
				
				$worksheet->write($a,103, $s->{o_retail},$border1); # outright sales retail
				$worksheet->write($a,104, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
					if ($s->{o_retail} <= 0){
					$worksheet->write($a,105, "",$subt); }
					else{
					$worksheet->write($a,105, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
					
				$worksheet->write($a,106, "",$border1);
				$worksheet->write($a,107, "",$subt);
				$worksheet->write($a,108, "",$border1);
				$worksheet->write($a,109, "",$subt);			
				
				$worksheet->write($a,110, $s->{o_retail},$border1);
				$worksheet->write($a,111, $s->{o_margin},$border1);
					if ($s->{c_retail} <= 0){
						$worksheet->write($a,112, "",$subt); }
					else{
						$worksheet->write($a,112, $s->{o_margin}/$s->{o_retail},$subt); }
						
				$worksheet->write($a,113, "",$border1);
				$worksheet->write($a,114, "",$subt);
				$worksheet->write($a,115, "",$border1);
				$worksheet->write($a,116, "",$subt);	
				
				$worksheet->write($a,117, $s->{amt501000},$border1);
				$worksheet->write($a,118, "",$border1);
				$worksheet->write($a,119, "",$border1);
				
				$worksheet->write($a,120, $s->{amt503200},$border1);
				$worksheet->write($a,121, "",$border1);
				$worksheet->write($a,122, "",$border1);
							
				$worksheet->write($a,123, $s->{amt503250},$border1);
				$worksheet->write($a,124, "",$border1);
				$worksheet->write($a,125, "",$border1);
					
				$worksheet->write($a,126, $s->{amt503500},$border1);
				$worksheet->write($a,127, "",$border1);
				$worksheet->write($a,128, "",$border1);
					
				$worksheet->write($a,129, $s->{amt506000},$border1);
				$worksheet->write($a,130, "",$border1);
				$worksheet->write($a,131, "",$border1);
				
				$worksheet->write($a,132, $s->{amt503000},$border1);
				$worksheet->write($a,133, "",$border1);
				$worksheet->write($a,134, "",$border1);
						
				$worksheet->write($a,135, $s->{amt507000},$border1);
				$worksheet->write($a,136, "",$border1);
				$worksheet->write($a,137, "",$border1);
				
				$worksheet->write($a,138, $s->{amt999998},$border1);
				$worksheet->write($a,139, "",$border1);
				$worksheet->write($a,140, "",$border1);
				
				$worksheet->write($a,141, $s->{amt505000},$border1);
				$worksheet->write($a,142, "",$border1);
				$worksheet->write($a,143, "",$border1);
							
				$worksheet->write($a,144, $s->{amt432000},$border1);
				$worksheet->write($a,145, "",$border1);
				$worksheet->write($a,146, "",$border1);
			
				$worksheet->write($a,147, $s->{amt433000},$border1);
				$worksheet->write($a,148, "",$border1);
				$worksheet->write($a,149, "",$border1);
				
				$worksheet->write($a,150, $s->{amt458490},$border1);
				$worksheet->write($a,151, "",$border1);
				$worksheet->write($a,152, "",$border1);
				
				$worksheet->write($a,153, $s->{amt504000},$border1);
				$worksheet->write($a,154, "",$border1);
				$worksheet->write($a,155, "",$border1);
				
				$worksheet->write($a,156, $s->{c_retail},$border1);
				$worksheet->write($a,157, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,158, "",$subt); }
					else{
						$worksheet->write($a,158, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
						
				$worksheet->write($a,159, "",$border1);
				$worksheet->write($a,160, "",$subt);
				$worksheet->write($a,161, "",$border1);
				$worksheet->write($a,162, "",$subt);	
				
				$worksheet->write($a,163, $s->{c_retail},$border1);
				$worksheet->write($a,164, $s->{c_margin},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,165, "",$subt); }
					else{
						$worksheet->write($a,165, $s->{c_margin}/$s->{c_retail},$subt); }
					
				$worksheet->write($a,166, "",$border1);
				$worksheet->write($a,167, "",$subt);
				$worksheet->write($a,168, "",$border1);
				$worksheet->write($a,169, "",$subt);	
						
				$worksheet->write($a,170, $s->{amt434000},$border1);
				$worksheet->write($a,171, "",$border1);
				$worksheet->write($a,172, "",$border1);
						
				$worksheet->write($a,173, $s->{amt458550},$border1);
				$worksheet->write($a,174, "",$border1);
				$worksheet->write($a,175, "",$border1);
				
				$worksheet->write($a,176, $s->{amt460100},$border1);
				$worksheet->write($a,177, "",$border1);
				$worksheet->write($a,178, "",$border1);
					
				$worksheet->write($a,179, $s->{amt460200},$border1);
				$worksheet->write($a,180, "",$border1);
				$worksheet->write($a,181, "",$border1);
						
				$worksheet->write($a,182, $s->{amt460300},$border1);
				$worksheet->write($a,183, "",$border1);
				$worksheet->write($a,184, "",$border1);
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
				$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
				$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
					if (($s->{o_retail}+$s->{c_retail}) <= 0){
						$worksheet->write($a,9, "",$subt); }
					else{
						$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
				$worksheet->write($a,10, "",$border1);
				$worksheet->write($a,11, "",$subt);
				$worksheet->write($a,12, "",$border1);
				$worksheet->write($a,13, "",$subt);
				
				$worksheet->write($a,14, $s->{o_retail},$border1); # outright sales retail
				$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
					if ($s->{o_retail} <= 0){
						$worksheet->write($a,16, "",$subt); }
					else{
						$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
				
				$worksheet->write($a,17, "",$border1);
				$worksheet->write($a,18, "",$subt);
				$worksheet->write($a,19, "",$border1);
				$worksheet->write($a,20, "",$subt);			
				
				$worksheet->write($a,21, $s->{o_retail},$border1);
				$worksheet->write($a,22, $s->{o_margin},$border1);
					if ($s->{c_retail} <= 0){
						$worksheet->write($a,23, "",$subt); }
					else{
						$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
						
				$worksheet->write($a,24, "",$border1);
				$worksheet->write($a,25, "",$subt);
				$worksheet->write($a,26, "",$border1);
				$worksheet->write($a,27, "",$subt);	
				
				$worksheet->write($a,28, $s->{amt501000},$border1);
				$worksheet->write($a,29, "",$border1);
				$worksheet->write($a,30, "",$border1);
				
				$worksheet->write($a,31, $s->{amt503200},$border1);
				$worksheet->write($a,32, "",$border1);
				$worksheet->write($a,33, "",$border1);
							
				$worksheet->write($a,34, $s->{amt503250},$border1);
				$worksheet->write($a,35, "",$border1);
				$worksheet->write($a,36, "",$border1);
					
				$worksheet->write($a,37, $s->{amt503500},$border1);
				$worksheet->write($a,38, "",$border1);
				$worksheet->write($a,39, "",$border1);
					
				$worksheet->write($a,40, $s->{amt506000},$border1);
				$worksheet->write($a,41, "",$border1);
				$worksheet->write($a,42, "",$border1);
				
				$worksheet->write($a,43, $s->{amt503000},$border1);
				$worksheet->write($a,44, "",$border1);
				$worksheet->write($a,45, "",$border1);
							
				$worksheet->write($a,46, $s->{amt507000},$border1);
				$worksheet->write($a,47, "",$border1);
				$worksheet->write($a,48, "",$border1);
				
				$worksheet->write($a,49, $s->{amt999998},$border1);
				$worksheet->write($a,50, "",$border1);
				$worksheet->write($a,51, "",$border1);
				
				$worksheet->write($a,52, $s->{amt505000},$border1);
				$worksheet->write($a,53, "",$border1);
				$worksheet->write($a,54, "",$border1);
							
				$worksheet->write($a,55, $s->{amt432000},$border1);
				$worksheet->write($a,56, "",$border1);
				$worksheet->write($a,57, "",$border1);
				
				$worksheet->write($a,58, $s->{amt433000},$border1);
				$worksheet->write($a,59, "",$border1);
				$worksheet->write($a,60, "",$border1);
				
				$worksheet->write($a,61, $s->{amt458490},$border1);
				$worksheet->write($a,62, "",$border1);
				$worksheet->write($a,63, "",$border1);
				
				$worksheet->write($a,64, $s->{amt504000},$border1);
				$worksheet->write($a,65, "",$border1);
				$worksheet->write($a,66, "",$border1);
				
				$worksheet->write($a,67, $s->{c_retail},$border1);
				$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,69, "",$subt); }
					else{
						$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
						
				$worksheet->write($a,70, "",$border1);
				$worksheet->write($a,71, "",$subt);
				$worksheet->write($a,72, "",$border1);
				$worksheet->write($a,73, "",$subt);	
				
				$worksheet->write($a,74, $s->{c_retail},$border1);
				$worksheet->write($a,75, $s->{c_margin},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,76, "",$subt); }
					else{
						$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
					
				$worksheet->write($a,77, "",$border1);
				$worksheet->write($a,78, "",$subt);
				$worksheet->write($a,79, "",$border1);
				$worksheet->write($a,80, "",$subt);	
						
				$worksheet->write($a,81, $s->{amt434000},$border1);
				$worksheet->write($a,82, "",$border1);
				$worksheet->write($a,83, "",$border1);
						
				$worksheet->write($a,84, $s->{amt458550},$border1);
				$worksheet->write($a,85, "",$border1);
				$worksheet->write($a,86, "",$border1);
				
				$worksheet->write($a,87, $s->{amt460100},$border1);
				$worksheet->write($a,88, "",$border1);
				$worksheet->write($a,89, "",$border1);
					
				$worksheet->write($a,90, $s->{amt460200},$border1);
				$worksheet->write($a,91, "",$border1);
				$worksheet->write($a,92, "",$border1);
						
				$worksheet->write($a,93, $s->{amt460300},$border1);
				$worksheet->write($a,94, "",$border1);
				$worksheet->write($a,95, "",$border1);

				$a -= 1;
				$counter -= 1;	
			}
			
			else{						
				$worksheet->write($a,96, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
				$worksheet->write($a,97, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
					if (($s->{o_retail}+$s->{c_retail}) <= 0){
						$worksheet->write($a,98, "",$subt); }
					else{
						$worksheet->write($a,98, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
					
				$worksheet->write($a,99, "",$border1);
				$worksheet->write($a,100, "",$subt);
				$worksheet->write($a,101, "",$border1);
				$worksheet->write($a,102, "",$subt);
				
				$worksheet->write($a,103, $s->{o_retail},$border1); # outright sales retail
				$worksheet->write($a,104, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000},$border1); #outright margin
					if ($s->{o_retail} <= 0){
					$worksheet->write($a,105, "",$subt); }
					else{
					$worksheet->write($a,105, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000})/$s->{o_retail},$subt); }
					
				$worksheet->write($a,106, "",$border1);
				$worksheet->write($a,107, "",$subt);
				$worksheet->write($a,108, "",$border1);
				$worksheet->write($a,109, "",$subt);			
				
				$worksheet->write($a,110, $s->{o_retail},$border1);
				$worksheet->write($a,111, $s->{o_margin},$border1);
					if ($s->{c_retail} <= 0){
						$worksheet->write($a,112, "",$subt); }
					else{
						$worksheet->write($a,112, $s->{o_margin}/$s->{o_retail},$subt); }
						
				$worksheet->write($a,113, "",$border1);
				$worksheet->write($a,114, "",$subt);
				$worksheet->write($a,115, "",$border1);
				$worksheet->write($a,116, "",$subt);	
				
				$worksheet->write($a,117, $s->{amt501000},$border1);
				$worksheet->write($a,118, "",$border1);
				$worksheet->write($a,119, "",$border1);
				
				$worksheet->write($a,120, $s->{amt503200},$border1);
				$worksheet->write($a,121, "",$border1);
				$worksheet->write($a,122, "",$border1);
							
				$worksheet->write($a,123, $s->{amt503250},$border1);
				$worksheet->write($a,124, "",$border1);
				$worksheet->write($a,125, "",$border1);
					
				$worksheet->write($a,126, $s->{amt503500},$border1);
				$worksheet->write($a,127, "",$border1);
				$worksheet->write($a,128, "",$border1);
					
				$worksheet->write($a,129, $s->{amt506000},$border1);
				$worksheet->write($a,130, "",$border1);
				$worksheet->write($a,131, "",$border1);
				
				$worksheet->write($a,132, $s->{amt503000},$border1);
				$worksheet->write($a,133, "",$border1);
				$worksheet->write($a,134, "",$border1);
						
				$worksheet->write($a,135, $s->{amt507000},$border1);
				$worksheet->write($a,136, "",$border1);
				$worksheet->write($a,137, "",$border1);
				
				$worksheet->write($a,138, $s->{amt999998},$border1);
				$worksheet->write($a,139, "",$border1);
				$worksheet->write($a,140, "",$border1);
				
				$worksheet->write($a,141, $s->{amt505000},$border1);
				$worksheet->write($a,142, "",$border1);
				$worksheet->write($a,143, "",$border1);
							
				$worksheet->write($a,144, $s->{amt432000},$border1);
				$worksheet->write($a,145, "",$border1);
				$worksheet->write($a,146, "",$border1);
			
				$worksheet->write($a,147, $s->{amt433000},$border1);
				$worksheet->write($a,148, "",$border1);
				$worksheet->write($a,149, "",$border1);
				
				$worksheet->write($a,150, $s->{amt458490},$border1);
				$worksheet->write($a,151, "",$border1);
				$worksheet->write($a,152, "",$border1);
				
				$worksheet->write($a,153, $s->{amt504000},$border1);
				$worksheet->write($a,154, "",$border1);
				$worksheet->write($a,155, "",$border1);
				
				$worksheet->write($a,156, $s->{c_retail},$border1);
				$worksheet->write($a,157, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,158, "",$subt); }
					else{
						$worksheet->write($a,158, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
						
				$worksheet->write($a,159, "",$border1);
				$worksheet->write($a,160, "",$subt);
				$worksheet->write($a,161, "",$border1);
				$worksheet->write($a,162, "",$subt);	
				
				$worksheet->write($a,163, $s->{c_retail},$border1);
				$worksheet->write($a,164, $s->{c_margin},$border1);
					if ($s->{c_retail} le 0){
						$worksheet->write($a,165, "",$subt); }
					else{
						$worksheet->write($a,165, $s->{c_margin}/$s->{c_retail},$subt); }
					
				$worksheet->write($a,166, "",$border1);
				$worksheet->write($a,167, "",$subt);
				$worksheet->write($a,168, "",$border1);
				$worksheet->write($a,169, "",$subt);	
						
				$worksheet->write($a,170, $s->{amt434000},$border1);
				$worksheet->write($a,171, "",$border1);
				$worksheet->write($a,172, "",$border1);
						
				$worksheet->write($a,173, $s->{amt458550},$border1);
				$worksheet->write($a,174, "",$border1);
				$worksheet->write($a,175, "",$border1);
				
				$worksheet->write($a,176, $s->{amt460100},$border1);
				$worksheet->write($a,177, "",$border1);
				$worksheet->write($a,178, "",$border1);
					
				$worksheet->write($a,179, $s->{amt460200},$border1);
				$worksheet->write($a,180, "",$border1);
				$worksheet->write($a,181, "",$border1);
						
				$worksheet->write($a,182, $s->{amt460300},$border1);
				$worksheet->write($a,183, "",$border1);
				$worksheet->write($a,184, "",$border1);
				
					if ($s->{merch_group_code_rev} eq 'SU' and ($s->{store_code} eq '2001' or $s->{store_code} eq '2001W')) {
						$s1f2_o_retail += $s->{o_retail};
						$s1f2_o_margin += $s->{o_margin};
						$s1f2_c_retail += $s->{c_retail};
						$s1f2_c_margin += $s->{c_margin};
						$s1f2_amt501000 += $s->{amt501000};
						$s1f2_amt503200 += $s->{amt503200};
						$s1f2_amt503250 += $s->{amt503250};
						$s1f2_amt503500 += $s->{amt503500};
						$s1f2_amt506000 += $s->{amt506000};
						$s1f2_amt503000 += $s->{amt503000};
						$s1f2_amt507000 += $s->{amt507000};
						$s1f2_amt999998 += $s->{amt999998};
						$s1f2_amt504000 += $s->{amt504000};
						$s1f2_amt505000 += $s->{amt505000};
						$s1f2_amt432000 += $s->{amt432000};
						$s1f2_amt433000 += $s->{amt433000};
						$s1f2_amt458490 += $s->{amt458490};
						$s1f2_amt434000 += $s->{amt434000};
						$s1f2_amt458550 += $s->{amt458550};
						$s1f2_amt460100 += $s->{amt460100};
						$s1f2_amt460200 += $s->{amt460200};
						$s1f2_amt460300 += $s->{amt460300};
						
						$s1f2_counter ++; # once value = 2, we'll have a summation of s1 and f2						
					}
			}		
			
			$a++;
			$counter++;			
		}
		
		if($s1f2_counter eq 2){
			$worksheet->write($a,5, "",$desc);
			$worksheet->write($a,6, "METRO COLON + F2",$desc);
			
			$worksheet->write($a,96, $s1f2_o_retail+$s1f2_c_retail,$border1);
			$worksheet->write($a,97, $s1f2_o_margin+$s1f2_c_margin+$s1f2_amt501000+$s1f2_amt432000+$s1f2_amt433000+$s1f2_amt458490+$s1f2_amt503200+$s1f2_amt503250+$s1f2_amt503500+$ss1f2_amt506000+$s1f2_amt503000+$s1f2_amt507000+$s1f2_amt999998+$s1f2_amt505000+$s1f2_amt504000+$s1f2_amt434000+$s1f2_amt458550+$s1f2_amt460100+$s1f2_amt460200+$s1f2_amt460300,$border1);
				if (($s1f2_o_retail+$s1f2_c_retail) <= 0){
					$worksheet->write($a,98, "",$subt); }
				else{
					$worksheet->write($a,98, ($s1f2_o_margin+$s1f2_c_margin+$s1f2_amt501000+$s1f2_amt432000+$s1f2_amt433000+$s1f2_amt458490+$s1f2_amt503200+$s1f2_amt503250+$s1f2_amt503500+$ss1f2_amt506000+$s1f2_amt503000+$s1f2_amt507000+$s1f2_amt999998+$s1f2_amt505000+$s1f2_amt504000+$s1f2_amt434000+$s1f2_amt458550+$s1f2_amt460100+$s1f2_amt460200+$s1f2_amt460300)/($s1f2_o_retail+$s1f2_c_retail),$subt); }
			
			$worksheet->write($a,99, "",$border1);
			$worksheet->write($a,100, "",$subt);
			$worksheet->write($a,101, "",$border1);
			$worksheet->write($a,102, "",$subt);
			
			$worksheet->write($a,103, $s1f2_o_retail,$border1);
			$worksheet->write($a,104, $s1f2_o_margin+$s1f2_amt501000+$s1f2_amt432000+$s1f2_amt433000+$s1f2_amt458490+$s1f2_amt503200+$s1f2_amt503250+$s1f2_amt503500+$ss1f2_amt506000+$s1f2_amt503000+$s1f2_amt507000+$s1f2_amt999998+$s1f2_amt505000+$s1f2_amt504000,$border1);
				if ($s1f2_o_retail <= 0){
					$worksheet->write($a,105, "",$subt); }
				else{
					$worksheet->write($a,105, ($s1f2_o_margin+$s1f2_amt501000+$s1f2_amt432000+$s1f2_amt433000+$s1f2_amt458490+$s1f2_amt503200+$s1f2_amt503250+$s1f2_amt503500+$ss1f2_amt506000+$s1f2_amt503000+$s1f2_amt507000+$s1f2_amt999998+$s1f2_amt505000+$s1f2_amt504000)/$s1f2_o_retail,$subt); }
			
			$worksheet->write($a,106, "",$border1);
			$worksheet->write($a,107, "",$subt);
			$worksheet->write($a,108, "",$border1);
			$worksheet->write($a,109, "",$subt);			
			
			$worksheet->write($a,110, $s1f2_o_retail,$border1);
			$worksheet->write($a,111, $s1f2_o_margin,$border1);
				if ($s1f2_c_retail <= 0){
					$worksheet->write($a,112, "",$subt); }
				else{
					$worksheet->write($a,112, $s1f2_o_margin/$s1f2_o_retail,$subt); }
					
			$worksheet->write($a,113, "",$border1);
			$worksheet->write($a,114, "",$subt);
			$worksheet->write($a,115, "",$border1);
			$worksheet->write($a,116, "",$subt);	
			
			$worksheet->write($a,117, $s1f2_amt501000,$border1);
			$worksheet->write($a,118, "",$border1);
			$worksheet->write($a,119, "",$border1);
			
			$worksheet->write($a,120, $s1f2_amt503200,$border1);
			$worksheet->write($a,121, "",$border1);
			$worksheet->write($a,122, "",$border1);
						
			$worksheet->write($a,123, $s1f2_amt503250,$border1);
			$worksheet->write($a,124, "",$border1);
			$worksheet->write($a,125, "",$border1);
				
			$worksheet->write($a,126, $s1f2_amt503500,$border1);
			$worksheet->write($a,127, "",$border1);
			$worksheet->write($a,128, "",$border1);
				
			$worksheet->write($a,129, $s1f2_amt506000,$border1);
			$worksheet->write($a,130, "",$border1);
			$worksheet->write($a,131, "",$border1);
			
			$worksheet->write($a,132, $s1f2_amt503000,$border1);
			$worksheet->write($a,133, "",$border1);
			$worksheet->write($a,134, "",$border1);
						
			$worksheet->write($a,135, $s1f2_amt507000,$border1);
			$worksheet->write($a,136, "",$border1);
			$worksheet->write($a,137, "",$border1);
			
			$worksheet->write($a,138, $s1f2_amt999998,$border1);
			$worksheet->write($a,139, "",$border1);
			$worksheet->write($a,140, "",$border1);
			
			$worksheet->write($a,141, $s1f2_amt505000,$border1);
			$worksheet->write($a,142, "",$border1);
			$worksheet->write($a,143, "",$border1);
						
			$worksheet->write($a,144, $s1f2_amt432000,$border1);
			$worksheet->write($a,145, "",$border1);
			$worksheet->write($a,146, "",$border1);
		
			$worksheet->write($a,147, $s1f2_amt433000,$border1);
			$worksheet->write($a,148, "",$border1);
			$worksheet->write($a,149, "",$border1);
			
			$worksheet->write($a,150, $s1f2_amt458490,$border1);
			$worksheet->write($a,151, "",$border1);
			$worksheet->write($a,152, "",$border1);
			
			$worksheet->write($a,153, $s1f2_amt504000,$border1);
			$worksheet->write($a,154, "",$border1);
			$worksheet->write($a,155, "",$border1);
			
			$worksheet->write($a,156, $s1f2_c_retail,$border1);
			$worksheet->write($a,157, $s1f2_c_margin+$s1f2_amt434000+$s1f2_amt458550+$s1f2_amt460100+$s1f2_amt460200+$s1f2_amt460300,$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,158, "",$subt); }
				else{
					$worksheet->write($a,158, ($s1f2_c_margin+$s1f2_amt434000+$s1f2_amt458550+$s1f2_amt460100+$s1f2_amt460200+$s1f2_amt460300)/$s1f2_c_retail,$subt); }
					
			$worksheet->write($a,159, "",$border1);
			$worksheet->write($a,160, "",$subt);
			$worksheet->write($a,161, "",$border1);
			$worksheet->write($a,162, "",$subt);	
			
			$worksheet->write($a,163, $s1f2_c_retail,$border1);
			$worksheet->write($a,164, $s1f2_c_margin,$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,165, "",$subt); }
				else{
					$worksheet->write($a,165, $s1f2_c_margin/$s1f2_c_retail,$subt); }
				
			$worksheet->write($a,166, "",$border1);
			$worksheet->write($a,167, "",$subt);
			$worksheet->write($a,168, "",$border1);
			$worksheet->write($a,169, "",$subt);	
					
			$worksheet->write($a,170, $s1f2_amt434000,$border1);
			$worksheet->write($a,171, "",$border1);
			$worksheet->write($a,172, "",$border1);
					
			$worksheet->write($a,173, $s1f2_amt458550,$border1);
			$worksheet->write($a,174, "",$border1);
			$worksheet->write($a,175, "",$border1);
			
			$worksheet->write($a,176, $s1f2_amt460100,$border1);
			$worksheet->write($a,177, "",$border1);
			$worksheet->write($a,178, "",$border1);
				
			$worksheet->write($a,179, $s1f2_amt460200,$border1);
			$worksheet->write($a,180, "",$border1);
			$worksheet->write($a,181, "",$border1);
					
			$worksheet->write($a,182, $s1f2_amt460300,$border1);
			$worksheet->write($a,183, "",$border1);
			$worksheet->write($a,184, "",$border1);
			
			$worksheet->set_row( $a, undef, undef, 1, undef, undef ); #we hide this row
			
			$s1f2_row = $a;
			$s1f2_counter = 0;
			
			$a++;
			$counter++;
		}	
	}
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){
	
	$sls_2 = $dbh_csv->prepare (qq{SELECT store_code, store_description, merch_group_code_rev, SUM(otr_cost) AS o_cost, SUM(otr_retail) AS o_retail, SUM(otr_margin) AS o_margin, SUM(con_cost) AS c_cost, SUM(con_retail) AS c_retail, SUM(con_margin) AS c_margin, SUM(debit_amt432000) AS deb_amt432000, SUM(credit_amt432000) AS cre_amt432000, SUM(amount432000) AS amt432000, SUM(debit_amt433000) AS deb_amt433000, SUM(credit_amt433000) AS cre_amt433000, SUM(amount433000) AS amt433000, SUM(debit_amt458490) AS deb_amt458490, SUM(credit_amt458490) AS cre_amt458490, SUM(amount458490) AS amt458490, SUM(debit_amt434000) AS deb_amt434000, SUM(credit_amt434000) AS cre_amt434000, SUM(amount434000) AS amt434000, SUM(debit_amt458550) AS deb_amt458550, SUM(credit_amt458550) AS cre_amt458550, SUM(amount458550) AS amt458550, SUM(debit_amt460100) AS deb_amt460100, SUM(credit_amt460100) AS cre_amt460100, SUM(amount460100) AS amt460100, SUM(debit_amt460200) AS deb_amt460200, SUM(credit_amt460200) AS cre_amt460200, SUM(amount460200) AS amt460200, SUM(debit_amt460300) AS deb_amt460300, SUM(credit_amt460300) AS cre_amt460300, SUM(amount460300) AS amt460300, SUM(debit_amt503200) AS deb_amt503200, SUM(credit_amt503200) AS cre_amt503200, SUM(amount503200) AS amt503200 , SUM(debit_amt503250) AS deb_amt503250, SUM(credit_amt503250) AS cre_amt503250, SUM(amount503250) AS amt503250 , SUM(debit_amt503500) AS deb_amt503500, SUM(credit_amt503500) AS cre_amt503500, SUM(amount503500) AS amt503500 , SUM(debit_amt506000) AS deb_amt506000, SUM(credit_amt506000) AS cre_amt506000, SUM(amount506000) AS amt506000, SUM(amount501000) AS amt501000, SUM(amount503000) AS amt503000, SUM(amount507000) AS amt507000, SUM(amount999998) AS amt999998, SUM(amount505000) AS amt505000, SUM(amount504000) AS amt504000
								 FROM $table 
								 WHERE (merch_group_code_rev = 'DS' or merch_group_code_rev = 'SU') AND (((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
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
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40'))
								 GROUP BY store_code, store_description , merch_group_code_rev
								 ORDER BY store_code, merch_group_code_rev
								});
	$sls_2->execute();	

	while(my $s = $sls_2->fetchrow_hashref()){
	
	$worksheet->write($a,5, $s->{store_code},$desc);
	$worksheet->write($a,6, $s->{store_description},$desc);
	
		if($s->{merch_group_code_rev} eq 'DS'){
			$worksheet->write($a,7, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
			$worksheet->write($a,8, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
				if (($s->{o_retail}+$s->{c_retail}) <= 0){
					$worksheet->write($a,9, "",$subt); }
				else{
					$worksheet->write($a,9, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
			
			$worksheet->write($a,10, "",$border1);
			$worksheet->write($a,11, "",$subt);
			$worksheet->write($a,12, "",$border1);
			$worksheet->write($a,13, "",$subt);
			
			$worksheet->write($a,14, $s->{o_retail},$border1); # outright sales retail
			$worksheet->write($a,15, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
				if ($s->{o_retail} <= 0){
					$worksheet->write($a,16, "",$subt); }
				else{
					$worksheet->write($a,16, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
			
			$worksheet->write($a,17, "",$border1);
			$worksheet->write($a,18, "",$subt);
			$worksheet->write($a,19, "",$border1);
			$worksheet->write($a,20, "",$subt);			
			
			$worksheet->write($a,21, $s->{o_retail},$border1);
			$worksheet->write($a,22, $s->{o_margin},$border1);
				if ($s->{c_retail} <= 0){
					$worksheet->write($a,23, "",$subt); }
				else{
					$worksheet->write($a,23, $s->{o_margin}/$s->{o_retail},$subt); }
					
			$worksheet->write($a,24, "",$border1);
			$worksheet->write($a,25, "",$subt);
			$worksheet->write($a,26, "",$border1);
			$worksheet->write($a,27, "",$subt);	
			
			$worksheet->write($a,28, $s->{amt501000},$border1);
			$worksheet->write($a,29, "",$border1);
			$worksheet->write($a,30, "",$border1);
			
			$worksheet->write($a,31, $s->{amt503200},$border1);
			$worksheet->write($a,32, "",$border1);
			$worksheet->write($a,33, "",$border1);
						
			$worksheet->write($a,34, $s->{amt503250},$border1);
			$worksheet->write($a,35, "",$border1);
			$worksheet->write($a,36, "",$border1);
				
			$worksheet->write($a,37, $s->{amt503500},$border1);
			$worksheet->write($a,38, "",$border1);
			$worksheet->write($a,39, "",$border1);
				
			$worksheet->write($a,40, $s->{amt506000},$border1);
			$worksheet->write($a,41, "",$border1);
			$worksheet->write($a,42, "",$border1);
			
			$worksheet->write($a,43, $s->{amt503000},$border1);
			$worksheet->write($a,44, "",$border1);
			$worksheet->write($a,45, "",$border1);
						
			$worksheet->write($a,46, $s->{amt507000},$border1);
			$worksheet->write($a,47, "",$border1);
			$worksheet->write($a,48, "",$border1);
			
			$worksheet->write($a,49, $s->{amt999998},$border1);
			$worksheet->write($a,50, "",$border1);
			$worksheet->write($a,51, "",$border1);
			
			$worksheet->write($a,52, $s->{amt505000},$border1);
			$worksheet->write($a,53, "",$border1);
			$worksheet->write($a,54, "",$border1);
						
			$worksheet->write($a,55, $s->{amt432000},$border1);
			$worksheet->write($a,56, "",$border1);
			$worksheet->write($a,57, "",$border1);
			
			$worksheet->write($a,58, $s->{amt433000},$border1);
			$worksheet->write($a,59, "",$border1);
			$worksheet->write($a,60, "",$border1);
			
			$worksheet->write($a,61, $s->{amt458490},$border1);
			$worksheet->write($a,62, "",$border1);
			$worksheet->write($a,63, "",$border1);
			
			$worksheet->write($a,64, $s->{amt504000},$border1);
			$worksheet->write($a,65, "",$border1);
			$worksheet->write($a,66, "",$border1);
			
			$worksheet->write($a,67, $s->{c_retail},$border1);
			$worksheet->write($a,68, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,69, "",$subt); }
				else{
					$worksheet->write($a,69, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
					
			$worksheet->write($a,70, "",$border1);
			$worksheet->write($a,71, "",$subt);
			$worksheet->write($a,72, "",$border1);
			$worksheet->write($a,73, "",$subt);	
			
			$worksheet->write($a,74, $s->{c_retail},$border1);
			$worksheet->write($a,75, $s->{c_margin},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,76, "",$subt); }
				else{
					$worksheet->write($a,76, $s->{c_margin}/$s->{c_retail},$subt); }
				
			$worksheet->write($a,77, "",$border1);
			$worksheet->write($a,78, "",$subt);
			$worksheet->write($a,79, "",$border1);
			$worksheet->write($a,80, "",$subt);	
					
			$worksheet->write($a,81, $s->{amt434000},$border1);
			$worksheet->write($a,82, "",$border1);
			$worksheet->write($a,83, "",$border1);
						
			$worksheet->write($a,84, $s->{amt458550},$border1);
			$worksheet->write($a,85, "",$border1);
			$worksheet->write($a,86, "",$border1);
			
			$worksheet->write($a,87, $s->{amt460100},$border1);
			$worksheet->write($a,88, "",$border1);
			$worksheet->write($a,89, "",$border1);
				
			$worksheet->write($a,90, $s->{amt460200},$border1);
			$worksheet->write($a,91, "",$border1);
			$worksheet->write($a,92, "",$border1);
					
			$worksheet->write($a,93, $s->{amt460300},$border1);
			$worksheet->write($a,94, "",$border1);
			$worksheet->write($a,95, "",$border1);
		}
		
		else{	
			$worksheet->write($a,96, $s->{o_retail}+$s->{c_retail},$border1); # total sales retail
			$worksheet->write($a,97, $s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1); # total margin
				if (($s->{o_retail}+$s->{c_retail}) <= 0){
					$worksheet->write($a,98, "",$subt); }
				else{
					$worksheet->write($a,98, ($s->{o_margin}+$s->{c_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/($s->{o_retail}+$s->{c_retail}),$subt); }
				
			$worksheet->write($a,99, "",$border1);
			$worksheet->write($a,100, "",$subt);
			$worksheet->write($a,101, "",$border1);
			$worksheet->write($a,102, "",$subt);
			
			$worksheet->write($a,103, $s->{o_retail},$border1); # outright sales retail
			$worksheet->write($a,104, $s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000},$border1); #outright margin
				if ($s->{o_retail} <= 0){
				$worksheet->write($a,105, "",$subt); }
				else{
				$worksheet->write($a,105, ($s->{o_margin}+$s->{amt501000}+$s->{amt432000}+$s->{amt433000}+$s->{amt458490}+$s->{amt503200}+$s->{amt503250}+$s->{amt503500}+$s->{amt506000}+$s->{amt503000}+$s->{amt507000}+$s->{amt999998}+$s->{amt505000}+$s->{amt504000})/$s->{o_retail},$subt); }
				
			$worksheet->write($a,106, "",$border1);
			$worksheet->write($a,107, "",$subt);
			$worksheet->write($a,108, "",$border1);
			$worksheet->write($a,109, "",$subt);			
			
			$worksheet->write($a,110, $s->{o_retail},$border1);
			$worksheet->write($a,111, $s->{o_margin},$border1);
				if ($s->{c_retail} <= 0){
					$worksheet->write($a,112, "",$subt); }
				else{
					$worksheet->write($a,112, $s->{o_margin}/$s->{o_retail},$subt); }
					
			$worksheet->write($a,113, "",$border1);
			$worksheet->write($a,114, "",$subt);
			$worksheet->write($a,115, "",$border1);
			$worksheet->write($a,116, "",$subt);	
			
			$worksheet->write($a,117, $s->{amt501000},$border1);
			$worksheet->write($a,118, "",$border1);
			$worksheet->write($a,119, "",$border1);
			
			$worksheet->write($a,120, $s->{amt503200},$border1);
			$worksheet->write($a,121, "",$border1);
			$worksheet->write($a,122, "",$border1);
						
			$worksheet->write($a,123, $s->{amt503250},$border1);
			$worksheet->write($a,124, "",$border1);
			$worksheet->write($a,125, "",$border1);
				
			$worksheet->write($a,126, $s->{amt503500},$border1);
			$worksheet->write($a,127, "",$border1);
			$worksheet->write($a,128, "",$border1);
				
			$worksheet->write($a,129, $s->{amt506000},$border1);
			$worksheet->write($a,130, "",$border1);
			$worksheet->write($a,131, "",$border1);
			
			$worksheet->write($a,132, $s->{amt503000},$border1);
			$worksheet->write($a,133, "",$border1);
			$worksheet->write($a,134, "",$border1);
					
			$worksheet->write($a,135, $s->{amt507000},$border1);
			$worksheet->write($a,136, "",$border1);
			$worksheet->write($a,137, "",$border1);
			
			$worksheet->write($a,138, $s->{amt999998},$border1);
			$worksheet->write($a,139, "",$border1);
			$worksheet->write($a,140, "",$border1);
			
			$worksheet->write($a,141, $s->{amt505000},$border1);
			$worksheet->write($a,142, "",$border1);
			$worksheet->write($a,143, "",$border1);
						
			$worksheet->write($a,144, $s->{amt432000},$border1);
			$worksheet->write($a,145, "",$border1);
			$worksheet->write($a,146, "",$border1);
		
			$worksheet->write($a,147, $s->{amt433000},$border1);
			$worksheet->write($a,148, "",$border1);
			$worksheet->write($a,149, "",$border1);
			
			$worksheet->write($a,150, $s->{amt458490},$border1);
			$worksheet->write($a,151, "",$border1);
			$worksheet->write($a,152, "",$border1);
			
			$worksheet->write($a,153, $s->{amt504000},$border1);
			$worksheet->write($a,154, "",$border1);
			$worksheet->write($a,155, "",$border1);
			
			$worksheet->write($a,156, $s->{c_retail},$border1);
			$worksheet->write($a,157, $s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,158, "",$subt); }
				else{
					$worksheet->write($a,158, ($s->{c_margin}+$s->{amt434000}+$s->{amt458550}+$s->{amt460100}+$s->{amt460200}+$s->{amt460300})/$s->{c_retail},$subt); }
					
			$worksheet->write($a,159, "",$border1);
			$worksheet->write($a,160, "",$subt);
			$worksheet->write($a,161, "",$border1);
			$worksheet->write($a,162, "",$subt);	
			
			$worksheet->write($a,163, $s->{c_retail},$border1);
			$worksheet->write($a,164, $s->{c_margin},$border1);
				if ($s->{c_retail} le 0){
					$worksheet->write($a,165, "",$subt); }
				else{
					$worksheet->write($a,165, $s->{c_margin}/$s->{c_retail},$subt); }
				
			$worksheet->write($a,166, "",$border1);
			$worksheet->write($a,167, "",$subt);
			$worksheet->write($a,168, "",$border1);
			$worksheet->write($a,169, "",$subt);	
					
			$worksheet->write($a,170, $s->{amt434000},$border1);
			$worksheet->write($a,171, "",$border1);
			$worksheet->write($a,172, "",$border1);
					
			$worksheet->write($a,173, $s->{amt458550},$border1);
			$worksheet->write($a,174, "",$border1);
			$worksheet->write($a,175, "",$border1);
			
			$worksheet->write($a,176, $s->{amt460100},$border1);
			$worksheet->write($a,177, "",$border1);
			$worksheet->write($a,178, "",$border1);
				
			$worksheet->write($a,179, $s->{amt460200},$border1);
			$worksheet->write($a,180, "",$border1);
			$worksheet->write($a,181, "",$border1);
					
			$worksheet->write($a,182, $s->{amt460300},$border1);
			$worksheet->write($a,183, "",$border1);
			$worksheet->write($a,184, "",$border1);	
		
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
	foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94,96,97,99,101,103,104,106,108,110,111,113,115,117,118,120,121,123,124,126,127,129,130,132,133,135,136,138,139,141,142,144,145,147,148,150,151,153,154,156,157,159,161,163,164,166,168,170,171,173,174,176,177,179,180,182,183
 ){
		if($s1f2_row ne 0){
			my $sum1 = '=(SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . '))-'. xl_rowcol_to_cell( $s1f2_row, $col );
				$worksheet->write( $a, $col, $sum1, $bodyNum );	}
		else{
			my $sum1 = '=SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . ')';
				$worksheet->write( $a, $col, $sum1, $bodyNum ); }

			if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75 or $col eq 97 or $col eq 104 or $col eq 111 or $col eq 157 or $col eq 164){
				my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
			elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77 or $col eq 99 or $col eq 106 or $col eq 113 or $col eq 159 or $col eq 168){
				my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
			elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79 or $col eq 101 or $col eq 108 or $col eq 115 or $col eq 161 or $col eq 168){
				my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
			elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94 
			  or $col eq 118 or $col eq 121 or $col eq 124 or $col eq 127 or $col eq 130 or $col eq 133 or $col eq 136 or $col eq 139 or $col eq 142 or $col eq 145 or $col eq 148 or $col eq 151 or $col eq 154 or $col eq 171 or $col eq 174 or $col eq 177 or $col eq 180 or $col eq 183){
				my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
					$worksheet->write( $a, $col+1, $var, $bodyPct );	} 	}	}

else{
	foreach my $col( 7,8,10,12,14,15,17,19,21,22,24,26,28,29,31,32,34,35,37,38,40,41,43,44,46,47,49,50,52,53,55,56,58,59,61,62,64,65,67,68,70,72,74,75,77,79,81,82,84,85,87,88,90,91,93,94 ){
		if($s1f2_row ne 0){
			my $sum = '=(SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . '))-'. xl_rowcol_to_cell( $s1f2_row, $col );
			$worksheet->write( $a, $col, $sum, $bodyNum );}
		
		else{
			my $sum = '=SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . ')';
				$worksheet->write( $a, $col, $sum, $bodyNum );	}
			
			if ($col eq 8 or $col eq 15 or $col eq 22 or $col eq 68 or $col eq 75){
				my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
			elsif ($col eq 10 or $col eq 17 or $col eq 24 or $col eq 70 or $col eq 77){
				my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-3 ) . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
			elsif ($col eq 12 or $col eq 19 or $col eq 26 or $col eq 72 or $col eq 79){
				my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-5 ) . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );} 
			elsif ($col eq 29 or $col eq 32 or $col eq 35 or $col eq 38 or $col eq 41 or $col eq 44 or $col eq 47 or $col eq 50 or $col eq 53 or $col eq 56 or $col eq 59 or $col eq 62 or $col eq 65 or $col eq 82 or $col eq 85 or $col eq 88 or $col eq 91 or $col eq 94){
				my $var = '=('. xl_rowcol_to_cell( $a, $col ). '-' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
					$worksheet->write( $a, $col+1, $var, $bodyPct );	}	}	}	}


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
 open my $fh, ">", "consolidated_margin_x.csv" or die "consolidated_margin_x.csv: $!";

$test = '
SELECT 
SGD.STORE_FORMAT,
CASE WHEN TO_CHAR(SGD.STORE) = \'4002\' THEN \'2001W\' ELSE TO_CHAR(SGD.STORE) END AS STORE_CODE,
CASE WHEN SGD.STORE IN (\'2012\', \'2013\', \'3009\', \'4004\', \'3010\', \'3011\') THEN \'SU\' || SGD.STORE	 WHEN SGD.STORE IN (\'4002\') THEN \'SU2001W\'     WHEN SGD.STORE = \'2223\' THEN \'DS\' || SGD.STORE	 ELSE SGD.MERCH_GROUP_CODE || SGD.STORE END AS STORE,
SGD.STORE_NAME STORE_DESCRIPTION, 
CASE WHEN (SGD.MERCH_GROUP_CODE = \'OT\' AND SGD.DIVISION = 9000) THEN \'DS\'     WHEN (SGD.MERCH_GROUP_CODE = \'OT\' AND SGD.DIVISION = 8500) THEN \'SU\'     WHEN (SGD.MERCH_GROUP_CODE = \'OT\' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO = 8040) THEN \'SU\'     WHEN (SGD.MERCH_GROUP_CODE = \'OT\' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO != 8040) THEN \'DS\' ELSE SGD.MERCH_GROUP_CODE END AS MERCH_GROUP_CODE_REV,
SGD.MERCH_GROUP_CODE, 
SGD.MERCH_GROUP_DESC, 
SGD.ATTRIB1 GROUP_CODE, 
SGD.ATTRIB2 GROUP_DESC, 
SGD.DIVISION, 
SGD.DIV_NAME DIVISION_DESC, 
SGD.GROUP_NO DEPARTMENT_CODE, 
SGD.GROUP_NAME DEPARTMENT_DESC, 
(SUM(OTR.OTR_COST)/1000) OTR_COST, (SUM(OTR.OTR_RETAIL)/1000) OTR_RETAIL, (SUM(OTR.OTR_MARGIN)/1000) OTR_MARGIN, 
(SUM(CON.CON_COST)/1000) CON_COST, (SUM(CON.CON_RETAIL)/1000) CON_RETAIL, (SUM(CON.CON_MARGIN)/1000) CON_MARGIN,
(SUM(OTR501000.DEBIT_AMT501000)/1000) DEBIT_AMT501000, (SUM(OTR501000.CREDIT_AMT501000)/1000) CREDIT_AMT501000, (SUM(OTR501000.AMOUNT501000)/1000) AMOUNT501000,
(SUM(OTR503200.DEBIT_AMT503200)/1000) DEBIT_AMT503200, (SUM(OTR503200.CREDIT_AMT503200)/1000) CREDIT_AMT503200, (SUM(OTR503200.AMOUNT503200)/1000) AMOUNT503200,
(SUM(OTR503250.DEBIT_AMT503250)/1000) DEBIT_AMT503250, (SUM(OTR503250.CREDIT_AMT503250)/1000) CREDIT_AMT503250, (SUM(OTR503250.AMOUNT503250)/1000) AMOUNT503250,
(SUM(OTR503500.DEBIT_AMT503500)/1000) DEBIT_AMT503500, (SUM(OTR503500.CREDIT_AMT503500)/1000) CREDIT_AMT503500, (SUM(OTR503500.AMOUNT503500)/1000) AMOUNT503500,
(SUM(OTR506000.DEBIT_AMT506000)/1000) DEBIT_AMT506000, (SUM(OTR506000.CREDIT_AMT506000)/1000) CREDIT_AMT506000, (SUM(OTR506000.AMOUNT506000)/1000) AMOUNT506000,
(SUM(OTR503000.DEBIT_AMT503000)/1000) DEBIT_AMT503000, (SUM(OTR503000.CREDIT_AMT503000)/1000) CREDIT_AMT503000, (SUM(OTR503000.AMOUNT503000)/1000) AMOUNT503000,
(SUM(OTR507000.DEBIT_AMT507000)/1000) DEBIT_AMT507000, (SUM(OTR507000.CREDIT_AMT507000)/1000) CREDIT_AMT507000, (SUM(OTR507000.AMOUNT507000)/1000) AMOUNT507000,
(SUM(OTR999998.DEBIT_AMT999998)/1000) DEBIT_AMT999998, (SUM(OTR999998.CREDIT_AMT999998)/1000) CREDIT_AMT999998, (SUM(OTR999998.AMOUNT999998)/1000) AMOUNT999998,
(SUM(OTR505000.DEBIT_AMT505000)/1000) DEBIT_AMT505000, (SUM(OTR505000.CREDIT_AMT505000)/1000) CREDIT_AMT505000, (SUM(OTR505000.AMOUNT505000)/1000) AMOUNT505000,
(SUM(OTR504000.DEBIT_AMT504000)/1000) DEBIT_AMT504000, (SUM(OTR504000.CREDIT_AMT504000)/1000) CREDIT_AMT504000, (SUM(OTR504000.AMOUNT504000)/1000) AMOUNT504000,
(SUM(OTR432000.DEBIT_AMT432000)/1000) DEBIT_AMT432000, (SUM(OTR432000.CREDIT_AMT432000)/1000) CREDIT_AMT432000, (SUM(OTR432000.AMOUNT432000)/1000) AMOUNT432000,
(SUM(OTR433000.DEBIT_AMT433000)/1000) DEBIT_AMT433000, (SUM(OTR433000.CREDIT_AMT433000)/1000) CREDIT_AMT433000, (SUM(OTR433000.AMOUNT433000)/1000) AMOUNT433000,
(SUM(OTR458490.DEBIT_AMT458490)/1000) DEBIT_AMT458490, (SUM(OTR458490.CREDIT_AMT458490)/1000) CREDIT_AMT458490, (SUM(OTR458490.AMOUNT458490)/1000) AMOUNT458490,
(SUM(CON434000.DEBIT_AMT434000)/1000) DEBIT_AMT434000, (SUM(CON434000.CREDIT_AMT434000)/1000) CREDIT_AMT434000, (SUM(CON434000.AMOUNT434000)/1000) AMOUNT434000,
(SUM(CON458550.DEBIT_AMT458550)/1000) DEBIT_AMT458550, (SUM(CON458550.CREDIT_AMT458550)/1000) CREDIT_AMT458550, (SUM(CON458550.AMOUNT458550)/1000) AMOUNT458550,
(SUM(CON460100.DEBIT_AMT460100)/1000) DEBIT_AMT460100, (SUM(CON460100.CREDIT_AMT460100)/1000) CREDIT_AMT460100, (SUM(CON460100.AMOUNT460100)/1000) AMOUNT460100,
(SUM(CON460200.DEBIT_AMT460200)/1000) DEBIT_AMT460200, (SUM(CON460200.CREDIT_AMT460200)/1000) CREDIT_AMT460200, (SUM(CON460200.AMOUNT460200)/1000) AMOUNT460200,
(SUM(CON460300.DEBIT_AMT460300)/1000) DEBIT_AMT460300, (SUM(CON460300.CREDIT_AMT460300)/1000) CREDIT_AMT460300, (SUM(CON460300.AMOUNT460300)/1000) AMOUNT460300,
CASE WHEN SGD.STORE IN (\'2009\',\'2012\',\'7176\',\'7003\',\'2006\',\'7004\',\'2007\',\'6008\',\'7005\',\'2010\',\'7300\',\'7009\',\'5006\',\'5005\',\'5004\',\'5003\',\'5002\',\'5001\',\'7173\',\'3004\',\'3003\',\'3001\',\'4003\',\'7008\',\'7007\',\'7006\',\'3006\',\'3007\',\'2003\',\'7000\',\'4002\',\'3005\',\'3002\',\'2002\',\'2001\',\'2011\',\'2008\',\'7001\',\'2004\',\'7002\',\'2005\') THEN 0 ELSE 1 END AS NEW_FLG,
CASE WHEN SGD.STORE IN (\'2009\',\'2012\',\'7176\',\'7003\',\'2006\',\'7004\',\'2007\',\'6008\',\'7005\',\'2010\',\'7300\',\'7009\',\'5006\',\'5005\',\'5004\',\'5003\',\'5002\',\'5001\',\'7173\',\'3004\',\'3003\',\'3001\',\'4003\',\'7008\',\'7007\',\'7006\',\'3006\',\'3007\',\'2003\',\'7000\',\'4002\',\'3005\',\'3002\',\'2002\',\'2001\',\'2011\',\'2008\',\'7001\',\'2004\',\'7002\',\'2005\') THEN 1 ELSE 0 END AS MATURED_FLG
FROM
	
	(SELECT S.STORE_FORMAT, S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
		FROM
			(SELECT DISTINCT STORE_FORMAT, STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (6008) UNION ALL SELECT DISTINCT 8 AS STORE_FORMAT, WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
			(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME 
				FROM DEPS D 
				  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
				  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
				  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION
			 WHERE I.DIVISION NOT IN (4000)
				UNION ALL
				SELECT \'Z\' AS MERCH_GROUP_CODE, \'OTHERS\' AS MERCH_GROUP_DESC, \'OT\' AS ATTRIB1, \'OTHERS\' AS ATTRIB2, 9999 AS DIVISION, \'Others\' AS DIV_NAME, 0 AS GROUP_NO, \'Default\' AS GROUP_NAME FROM DUAL 
				UNION ALL 
				SELECT \'Z\' AS MERCH_GROUP_CODE, \'OTHERS\' AS MERCH_GROUP_DESC, \'OT\' AS ATTRIB1, \'OTHERS\' AS ATTRIB2, 9999 AS DIVISION, \'Others\' AS DIV_NAME, 999 AS GROUP_NO, \'Others\' AS GROUP_NAME FROM DUAL)M
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
			WHERE TDH.TRAN_DATE = \'30-NOV-13\' AND TDH.TRAN_CODE = 1 AND DEPS.PURCHASE_TYPE = 0)T
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
			WHERE TDH.TRAN_DATE = \'30-NOV-13\' AND TDH.TRAN_CODE = 1 AND DEPS.PURCHASE_TYPE = 2)T
	GROUP BY T.LOCATION, T.GROUP_NO)CON
	ON SGD.STORE = CON.LOCATION AND SGD.GROUP_NO = CON.GROUP_NO
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT501000, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT501000,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT501000
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'501000\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' AND S.JE_SOURCE_NAME = \'Manual\'
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR501000
	ON SGD.STORE = OTR501000.STORE AND SGD.GROUP_NO = OTR501000.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT503200, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT503200,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT503200
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'503200\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR503200
	ON SGD.STORE = OTR503200.STORE AND SGD.GROUP_NO = OTR503200.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT503250, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT503250,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT503250
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'503250\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR503250
	ON SGD.STORE = OTR503250.STORE AND SGD.GROUP_NO = OTR503250.DEPARTMENT
	
	LEFT JOIN
	(SELECT MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT503500, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT503500,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT503500
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'503500\') AND CC.SEGMENT1=\'240\' --AND H.PERIOD_NAME=\'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR503500
	ON SGD.STORE = OTR503500.STORE AND SGD.GROUP_NO = OTR503500.DEPARTMENT
	
	LEFT JOIN
	(SELECT MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT506000, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT506000,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT506000
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'506000\') AND CC.SEGMENT1=\'240\' --AND H.PERIOD_NAME=\'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR506000
	ON SGD.STORE = OTR506000.STORE AND SGD.GROUP_NO = OTR506000.DEPARTMENT
	
	LEFT JOIN
	(SELECT MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT503000, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT503000,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT503000
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'503000\') AND CC.SEGMENT1=\'240\' --AND H.PERIOD_NAME=\'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR503000
	ON SGD.STORE = OTR503000.STORE AND SGD.GROUP_NO = OTR503000.DEPARTMENT
	
	LEFT JOIN
	(SELECT MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT507000, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT507000,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT507000
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'507000\') AND CC.SEGMENT1=\'240\' --AND H.PERIOD_NAME=\'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR507000
	ON SGD.STORE = OTR507000.STORE AND SGD.GROUP_NO = OTR507000.DEPARTMENT
	
	LEFT JOIN
	(SELECT MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT999998, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT999998,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT999998
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'999998\') AND CC.SEGMENT1=\'240\' --AND H.PERIOD_NAME=\'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR999998
	ON SGD.STORE = OTR999998.STORE AND SGD.GROUP_NO = OTR999998.DEPARTMENT
	
	LEFT JOIN
	(SELECT MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT505000, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT505000,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT505000
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'505000\') AND CC.SEGMENT1=\'240\' --AND H.PERIOD_NAME=\'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR505000
	ON SGD.STORE = OTR505000.STORE AND SGD.GROUP_NO = OTR505000.DEPARTMENT
	
	LEFT JOIN
	(SELECT MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT504000, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT504000,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT504000
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'504000\') AND CC.SEGMENT1=\'240\' --AND H.PERIOD_NAME=\'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR504000
	ON SGD.STORE = OTR504000.STORE AND SGD.GROUP_NO = OTR504000.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT432000, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT432000,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT432000
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
	    LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'432000\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR432000
	ON SGD.STORE = OTR432000.STORE AND SGD.GROUP_NO = OTR432000.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT433000, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT433000,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT433000
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
		LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'433000\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR433000
	ON SGD.STORE = OTR433000.STORE AND SGD.GROUP_NO = OTR433000.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT458490, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT458490,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT458490
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
		LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'458490\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)OTR458490
	ON SGD.STORE = OTR458490.STORE AND SGD.GROUP_NO = OTR458490.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT434000, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT434000,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT434000
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
		LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'434000\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)CON434000
	ON SGD.STORE = CON434000.STORE AND SGD.GROUP_NO = CON434000.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT458550, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT458550,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT458550
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
		LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'458550\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)CON458550
	ON SGD.STORE = CON458550.STORE AND SGD.GROUP_NO = CON458550.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT460100, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT460100,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT460100
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
		LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'460100\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)CON460100
	ON SGD.STORE = CON460100.STORE AND SGD.GROUP_NO = CON460100.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT460200, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT460200,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT460200
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
		LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'460200\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)CON460200
	ON SGD.STORE = CON460200.STORE AND SGD.GROUP_NO = CON460200.DEPARTMENT
	
	LEFT JOIN
	(SELECT CC.SEGMENT5 ACCOUNT, MG.RMS_LOC_NO STORE, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END AS DEPARTMENT, SUM(NVL(L.ACCOUNTED_DR,0)) DEBIT_AMT460300, SUM(NVL(L.ACCOUNTED_CR,0)) CREDIT_AMT460300,
	SUM(NVL(L.ACCOUNTED_CR,0)-NVL(L.ACCOUNTED_DR,0)) AMOUNT460300
	FROM GL_JE_HEADERS@RMS_FIN_DBLINK H 
		JOIN GL_JE_LINES@RMS_FIN_DBLINK L ON H.JE_HEADER_ID = L.JE_HEADER_ID 
		JOIN GL_JE_BATCHES@RMS_FIN_DBLINK B ON B.JE_BATCH_ID=H.JE_BATCH_ID
		JOIN GL_CODE_COMBINATIONS@RMS_FIN_DBLINK CC ON CC.CODE_COMBINATION_ID=L.CODE_COMBINATION_ID 
		JOIN GL_JE_SOURCES_TL@RMS_FIN_DBLINK S ON S.JE_SOURCE_NAME=H.JE_SOURCE
		JOIN FND_USER@RMS_FIN_DBLINK U ON U.USER_ID=H.CREATED_BY 
		LEFT JOIN MG_COMP_LOC_MAPPING MG ON CC.SEGMENT2 = MG.COA_SITE
		LEFT JOIN GROUPS G ON TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = G.GROUP_NO
	WHERE H.ACTUAL_FLAG=\'A\' AND CC.SEGMENT5 IN (\'460300\') AND CC.SEGMENT1 = \'240\' --AND H.PERIOD_NAME = \'NOV-14\' 
		AND H.ACCRUAL_REV_STATUS IS NULL AND H.STATUS=\'P\' 
		AND H.DEFAULT_EFFECTIVE_DATE = \'30-NOV-13\' --BETWEEN \'01-JAN-10\' AND \'31-DEC-10\'--AND H.POSTED_DATE BETWEEN \'06-APR-2014\' AND \'12-MAY-2014\'
	GROUP BY CC.SEGMENT5, MG.RMS_LOC_NO, CASE WHEN TRANSLATE(CC.SEGMENT4, \'1abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\', \'1\') = 0000 THEN 0 WHEN G.GROUP_NO IS NULL THEN 999 ELSE G.GROUP_NO END)CON460300
	ON SGD.STORE = CON460300.STORE AND SGD.GROUP_NO = CON460300.DEPARTMENT
	
GROUP BY 
SGD.STORE_FORMAT,
CASE WHEN TO_CHAR(SGD.STORE) = \'4002\' THEN \'2001W\' ELSE TO_CHAR(SGD.STORE) END,
CASE WHEN SGD.STORE IN (\'2012\', \'2013\', \'3009\', \'4004\', \'3010\', \'3011\') THEN \'SU\' || SGD.STORE 	 WHEN SGD.STORE IN (\'4002\') THEN \'SU2001W\'     WHEN SGD.STORE = \'2223\' THEN \'DS\' || SGD.STORE	 ELSE SGD.MERCH_GROUP_CODE || SGD.STORE END,
SGD.STORE_NAME, 
CASE WHEN (SGD.MERCH_GROUP_CODE = \'OT\' AND SGD.DIVISION = 9000) THEN \'DS\'     WHEN (SGD.MERCH_GROUP_CODE = \'OT\' AND SGD.DIVISION = 8500) THEN \'SU\'     WHEN (SGD.MERCH_GROUP_CODE = \'OT\' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO = 8040) THEN \'SU\'     WHEN (SGD.MERCH_GROUP_CODE = \'OT\' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO != 8040) THEN \'DS\' ELSE SGD.MERCH_GROUP_CODE END,
SGD.MERCH_GROUP_CODE, 
SGD.MERCH_GROUP_DESC, 
SGD.ATTRIB1, 
SGD.ATTRIB2, 
SGD.DIVISION, 
SGD.DIV_NAME, 
SGD.GROUP_NO, 
SGD.GROUP_NAME,
CASE WHEN SGD.STORE IN (\'2009\',\'2012\',\'7176\',\'7003\',\'2006\',\'7004\',\'2007\',\'6008\',\'7005\',\'2010\',\'7300\',\'7009\',\'5006\',\'5005\',\'5004\',\'5003\',\'5002\',\'5001\',\'7173\',\'3004\',\'3003\',\'3001\',\'4003\',\'7008\',\'7007\',\'7006\',\'3006\',\'3007\',\'2003\',\'7000\',\'4002\',\'3005\',\'3002\',\'2002\',\'2001\',\'2011\',\'2008\',\'7001\',\'2004\',\'7002\',\'2005\') THEN 0 ELSE 1 END,
CASE WHEN SGD.STORE IN (\'2009\',\'2012\',\'7176\',\'7003\',\'2006\',\'7004\',\'2007\',\'6008\',\'7005\',\'2010\',\'7300\',\'7009\',\'5006\',\'5005\',\'5004\',\'5003\',\'5002\',\'5001\',\'7173\',\'3004\',\'3003\',\'3001\',\'4003\',\'7008\',\'7007\',\'7006\',\'3006\',\'3007\',\'2003\',\'7000\',\'4002\',\'3005\',\'3002\',\'2002\',\'2001\',\'2011\',\'2008\',\'7001\',\'2004\',\'7002\',\'2005\') THEN 1 ELSE 0 END
ORDER BY 1, 3, 5, 7, 9';

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_lc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "consolidated_margin_x.csv: $!";
 
$sth->finish();
$dbh->disconnect;

}

# mailer
sub mail_grp1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1, $attachment_file_2 ) = @ARGV;

# $to = ' gerry.guanlao@metrogaisano.com, eric.redona@metrogaisano.com, lucille.malazarte@metrogaisano.com, tricia.luntao@metrogaisano.com, jj.moreno@metrogaisano.com, cj.jesena@metrogaisano.com, rex.cabanilla@metrogaisano.com ';

# $bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, christopher.calalang@metrogaisano.com, annalyn.conde@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com ';
		
$to = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com, rex.cabanilla@metrogaisano.com ';
# $bcc = 'kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, christopher.calalang@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com';
		
$subject = 'Sales and Margin Performance ' . $as_of;
$msgbody_file = 'message_BI.txt';

$attachment_file_1 = "CONSOLIDATED MARGIN - Summary (as of $as_of) v1.3x.xlsx";

my $msgbody = read_file( $msgbody_file );

my $attachment_data_1 = encode_base64( read_file( $attachment_file_1, 1 ));

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















