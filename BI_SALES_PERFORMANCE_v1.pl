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

$tst_query = $dbh->prepare($test_query);
$tst_query->execute();

while ( my $x =  $tst_query->fetchrow_hashref()){
	$test = $x->{STATUS};
} 

if ($test eq 1){

	$date = qq{ 
	SELECT WEEK_NUMBER_THIS_YEAR, DATE_KEY, TO_CHAR(DATE_FLD, 'DD Mon YYYY') DATE_FLD, WEEK_ST_DATE_KEY, WEEK_END_DATE_KEY 
	FROM DIM_DATE 
	WHERE DATE_FLD = (SELECT TO_DATE(VALUE,'YYYY-MM-DD') FROM ADMIN_ETL_SUMMARY)
	 };

	my $sth_date_1 = $dbh->prepare ($date);
	 $sth_date_1->execute;

	while (my $x = $sth_date_1->fetchrow_hashref()) {
		$wk_st_date_key = $x->{WEEK_ST_DATE_KEY};
		$wk_en_date_key = $x->{WEEK_END_DATE_KEY};
		$wk_number = $x->{WEEK_NUMBER_THIS_YEAR};
		$as_of = $x->{DATE_FLD};
	}

	$date_2 = qq{ 
	SELECT DATE_KEY1, TO_CHAR(DATE_FLD1, 'DD Mon YYYY') DATE_FLD1, TO_CHAR(DATE_FLD_LY1, 'DD Mon YYYY') DATE_FLD_LY1, DATE_KEY2, TO_CHAR(DATE_FLD2, 'DD Mon YYYY') DATE_FLD2, TO_CHAR(DATE_FLD_LY2, 'DD Mon YYYY') DATE_FLD_LY2, TO_CHAR(DATE_FLD_WTD_GREG, 'DD Mon YYYY') DATE_FLD_WTD_GREG_LY, DATE_KEY3, DATE_FLD3, MONTH_ST_DATE_KEY, MONTH_END_DATE_KEY FROM
		(SELECT DATE_KEY AS DATE_KEY1, DATE_FLD AS DATE_FLD1, DATE_FLD_LY AS DATE_FLD_LY1
		FROM DIM_DATE WHERE DATE_KEY = $wk_st_date_key),
		(SELECT DATE_KEY AS DATE_KEY2, DATE_FLD AS DATE_FLD2, DATE_FLD_LY AS DATE_FLD_LY2
		FROM DIM_DATE WHERE DATE_KEY = $wk_en_date_key),
		(SELECT DATE_KEY_LY AS DATE_KEY4, DATE_FLD_LY AS DATE_FLD_WTD_GREG
		FROM DIM_DATE WHERE TO_CHAR(DATE_FLD, 'DD Mon YYYY') = '$as_of'),
		(SELECT DATE_KEY AS DATE_KEY3, DATE_FLD AS DATE_FLD3, MONTH_ST_DATE_KEY, MONTH_END_DATE_KEY
		FROM DIM_DATE_PRL WHERE TO_CHAR(DATE_FLD, 'DD Mon YYYY') = '$as_of')
	 };

	my $sth_date_2 = $dbh->prepare ($date_2);
	 $sth_date_2->execute;
	 
	while (my $x = $sth_date_2->fetchrow_hashref()) {
		$wk_st_date_fld = $x->{DATE_FLD1};
		$wk_en_date_fld = $x->{DATE_FLD2};
		$wk_st_date_fld_ly = $x->{DATE_FLD_LY1};
		$wk_en_date_fld_ly = $x->{DATE_FLD_LY2};
		$wk_en_date_fld_greg_ly = $x->{DATE_FLD_WTD_GREG_LY};
		$mo_st_date_key = $x->{MONTH_ST_DATE_KEY};
		$mo_en_date_key = $x->{DATE_KEY3};
	}

	$date_3 = qq{ 
	SELECT DATE_KEY1, TO_CHAR(DATE_FLD1, 'DD Mon YYYY') DATE_FLD1, DATE_KEY_LY1, TO_CHAR(DATE_FLD_LY1, 'DD Mon YYYY') DATE_FLD_LY1, 
		   DATE_KEY2, TO_CHAR(DATE_FLD2, 'DD Mon YYYY') DATE_FLD2, DATE_KEY_LY2, TO_CHAR(DATE_FLD_LY2, 'DD Mon YYYY') DATE_FLD_LY2 FROM
		(SELECT DATE_KEY AS DATE_KEY1, DATE_FLD AS DATE_FLD1, DATE_KEY_LY AS DATE_KEY_LY1, DATE_FLD_LY AS DATE_FLD_LY1
		FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_st_date_key),
		(SELECT DATE_KEY AS DATE_KEY2, DATE_FLD AS DATE_FLD2, DATE_KEY_LY AS DATE_KEY_LY2, DATE_FLD_LY AS DATE_FLD_LY2
		FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_en_date_key)
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
	}
	
	#=============================== GROUP 1=================================================# 
 
	$workbook = Excel::Writer::XLSX->new("Daily Sales Performance - Summary (as of $as_of) v1.xlsx");
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

	&new_sheet($sheet = "Department");			
	&call_div;
	
	$workbook->close();
	
	my $pdf_job_1 = Win32::Job->new;
	$pdf_job_1->spawn( "cmd" , q{cmd /C java ecp_FileConverter "Daily Sales Performance - Summary (as of } . $as_of . q{) v1.xlsx" pdf});
	$pdf_job_1->run(60);	
	
	#&mail_grp1;	
	
	#=============================== GROUP 2=================================================#

	# $workbook = Excel::Writer::XLSX->new("Daily Sales Performance (as of $as_of) v1.xlsx");
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
	# $headN = $workbook->add_format( border => 1, align => 'center', valign => 'center', size => 11, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	# $headD = $workbook->add_format( border => 1, align => 'center', valign => 'center', size => 10, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	# $headDPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	# $headPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	# $headNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3, bg_color => $abo, bold => 1 );
	# $headNumber = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 3, bg_color => $abo, bold => 1 );
	# $head = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => '0.0 %', bg_color => $abo, bold => 1 );
	# $subhead = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => '0.0 %', bg_color => $ponkan, bold => 1 );
	# $bodyN = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10, bg_color => $sky, num_format => '0.0 %',  bold => 1);
	# $bodyPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => '0.0 %',  bold => 1);
	# $bodyNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 3,  bold => 1);
	# $body = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => '0.0 %',  bold => 1);
	# $subt = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => '0.0 %');
	# $down = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => '0.0 %', bg_color => $pula );

	# printf "Arc BI Sales Performance Part 2 \n";

	# &new_sheet($sheet = "Summary");
	# &call_str;

	# $workbook->close();
	
	# my $pdf_job_2 = Win32::Job->new;
	# $pdf_job_2->spawn( "cmd" , q{cmd /C java ecp_FileConverter "Daily Sales Performance (as of } . $as_of . q{) v1.xlsx" pdf});
	# $pdf_job_2->run(60);	

	# &mail_grp2;

	$tst_query->finish();
	$sth_date_1->finish();
	$sth_date_2->finish();
	$sth_date_3->finish();
	$dbh_csv->disconnect;
	$dbh->disconnect; 
	
	exit;
	
}

else{
	print "Test Status = ". $test ." \nETL still running\n";
	
	$tst_query->finish();
	
	sleep(900);
	
	goto START;
}
 
#================================= FUNCTIONS ==================================#

sub call_div {

$a = 10, $e = 10, $counter = 0, $total_wtd_net_ty = 0, $total_wtd_net_ly = 0, $total_wtd_target = 0, $total_mtd_net_ty = 0, $total_mtd_net_ly = 0 , $total_mtd_target = 0, $grp_wtd_net_ty = 0, $grp_wtd_net_ly = 0, $grp_wtd_target = 0, $grp_mtd_net_ty = 0, $grp_mtd_net_ly = 0 , $grp_mtd_target = 0;

$worksheet->write($a-9, 2, "Daily Sales Performance", $bold1);
$worksheet->write($a-8, 2, "WTD: $wk_st_date_fld - $mo_en_date_fld vs $wk_st_date_fld_ly - $wk_en_date_fld_greg_ly");
$worksheet->write($a-7, 2, "MTD: $mo_st_date_fld - $mo_en_date_fld vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-6, 2, "As of $as_of");	

##========================= COMP STORES ===========================##

&heading_2;
&heading;
&query_dept($new_flg = 0, $matured_flg = 1, $loc_desc = "COMP STORES");

##========================= ALL STORES ===========================##

$a += 7, $total_wtd_net_ty = 0, $total_wtd_net_ly = 0, $total_wtd_target = 0, $total_mtd_net_ty = 0, $total_mtd_net_ly = 0 , $total_mtd_target = 0, $grp_wtd_net_ty = 0, $grp_wtd_net_ly = 0, $grp_wtd_target = 0, $grp_mtd_net_ty = 0, $grp_mtd_net_ly = 0 , $grp_mtd_target = 0;

&heading_2;
&heading;
&query_dept($new_flg = 1, $matured_flg = 1, $loc_desc = "ALL STORES");

}

sub call_str {

$a=10, $d=0, $e=10, $counter=0, $comp_su = 0, $new_su = 0, $comp_nb = 0, $comp_hy = 0, $new_hy = 0, $comp_ds = 0, $new_ds = 0, $type_test = 1;

$worksheet->write($a-9, 3, "Daily Sales Performance", $bold1);
$worksheet->write($a-8, 3, "WTD: $wk_st_date_fld - $mo_en_date_fld vs $wk_st_date_fld_ly - $wk_en_date_fld_greg_ly");
$worksheet->write($a-7, 3, "MTD: $mo_st_date_fld - $mo_en_date_fld vs $mo_st_date_fld_ly - $mo_en_date_fld_ly");
$worksheet->write($a-6, 3, "As of $as_of");

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

#&calc9;

}


sub strComp_Su {

$div_name = "Comp";  $div_name3 = "Supermarket";
$mrch1 = 'SU'; $mrch2 = 'SU'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8040'; $dept_grp2 = '0000'; $dept_grp3 = '0000'; $dept_grp4 = '0000'; $dept_grp5 = '0000'; $dept_grp6 = '0000'; $dept_grp7 = '0000';

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU2001'; $store2 = 'SU2002'; $store3 = 'SU2003'; $store4 = 'SU2004'; $store5 = 'SU2006'; $store6 = 'SU2007'; $store7 = 'SU2009'; $store8 = 'SU2012'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2012'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 

	elsif($type_test eq 2){	
		
		&query_by_store;
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
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU2013'; $store2 = 'SU4004'; $store3 = 'SU3009'; $store4 = 'SU3010'; $store5 = 'SU3011'; $store6 = '0000'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2013'; $stor2 = '4004'; $stor3 = '3009'; $stor4 = '3010'; $stor5 = '3011'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 

	elsif($type_test eq 2) {	
		
		&query_by_store;
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_su=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 12, 13, 15 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_su,$col).','.xl_rowcol_to_cell($new_su,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 13){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '0', $bodyPct );
						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
						}
					}
					
					if ($col eq 10 or $col eq 15){
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
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU3000'; $store2 = 'SU3007'; $store3 = 'SU4003'; $store4 = 'DS3000'; $store5 = 'DS3007'; $store6 = 'DS4003'; $store7 = 'OT3000'; $store8 = 'OT3007'; $store9 = 'OT4003'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '3000'; $stor2 = '3007'; $stor3 = '4003'; $stor4 = '0000'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

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

		foreach my $col( 7, 8, 10, 12, 13, 15 ){
			my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).')';
				$worksheet->write( $a, $col, $sumTY, $bodyNum );
				
				if ($col eq 8 or $col eq 13){
					if( xl_rowcol_to_cell( $a, $col ) le 0){
						$worksheet->write( $a, $col+1, '0', $bodyPct );
					}
					else{
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
						$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
					}
				}
				
				if ($col eq 10 or $col eq 15){
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

sub strNew_Nb {

$div_name = "New"; $div_name2 = "Neighborhood";  $div_name3 = "Neighborhood Store";
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '0000'; $division_grp2 = '0000';  $division_grp3 = '0000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = '0000'; $store2 = '0000'; $store3 = '0000'; $store4 = '0000'; $store5 = '0000'; $store6 = '0000'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '0000'; $stor2 = '0000'; $stor3 = '0000'; $stor4 = '0000'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

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

		foreach my $col( 7, 8, 10, 12, 13, 15 ){
			my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).')';
				$worksheet->write( $a, $col, $sumTY, $bodyNum );
				
				if ($col eq 8 or $col eq 13){
					if( xl_rowcol_to_cell( $a, $col ) le 0){
						$worksheet->write( $a, $col+1, '0', $bodyPct );
					}
					else{
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
						$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
					}
				}
				
				if ($col eq 10 or $col eq 15){
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

sub strComp_Hy {

$div_name = "Comp";  $div_name3 = "Hypermarket";
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU2005'; $store2 = 'SU2008'; $store3 = 'SU2010'; $store4 = 'SU2011'; $store5 = 'DS2005'; $store6 = 'DS2008'; $store7 = 'DS2010'; $store8 = 'DS2011'; $store9 = 'DS2005'; $store10 = 'OT2005'; $store11 = 'OT2008'; $store12 = 'OT2010'; $store13 = 'OT2011'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2005'; $stor2 = '2008'; $stor3 = '2010'; $stor4 = '2011'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 

	elsif($type_test eq 2) {	

	&query_by_store;	
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
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU6001'; $store2 = 'SU6002'; $store3 = 'SU6003'; $store4 = 'SU6004'; $store5 = 'SU6005'; $store6 = 'SU6012'; $store7 = 'SU6009'; $store8 = 'SU6010'; $store9 = 'SU6011'; $store10 = 'DS6001'; $store11 = 'DS6002'; $store12 = 'DS6003'; $store13 = 'DS6004'; $store14 = 'DS6005'; $store15 = 'DS6012'; $store16 = 'DS6009'; $store17 = 'DS6010'; $store18 = 'DS6011'; $store19 = 'OT6002'; $store20 = 'OT6003'; $store21 = 'OT6004'; $store22 = 'OT6005'; $store23 = 'OT6012'; $store24 = 'OT6009'; $store25 = 'OT6010'; $store26 = 'OT6011';  $store27 = 'OT6000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '6001'; $stor2 = '6001'; $stor3 = '6003'; $stor4 = '6004'; $stor5 = '6005'; $stor6 = '6012'; $stor7 = '6009'; $stor8 = '6010'; $stor9 = '6011'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

	if($type_test eq 1){	
	
		&query_summary;	
		
		$counter = 4; 
		#&calc8; 
		#$a+=1; 
		
		#$worksheet->merge_range( $a-1, 4, $a-1, 6, 'Total', $bodyN );
		#$worksheet->merge_range( ($a)-$counter, 3, $a, 3, $div_name, $border2 );
		$counter = 0; $d=$a; 
		
	} 

	elsif($type_test eq 2) {	

	&query_by_store;	
	&calc8;

	$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
	$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
	$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

	$new_hy=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

		foreach my $col( 7, 8, 10, 12, 13, 15 ){
			my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_hy,$col).','.xl_rowcol_to_cell($new_hy,$col).')';
				$worksheet->write( $a, $col, $sumTY, $bodyNum );
				
				if ($col eq 8 or $col eq 13){
					if( xl_rowcol_to_cell( $a, $col ) le 0){
						$worksheet->write( $a, $col+1, '0', $bodyPct );
					}
					else{
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
						$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
					}
				}
				
				if ($col eq 10 or $col eq 15){
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

sub strComp_Ds {

$div_name = "Comp";  $div_name3 = "Department Store";
$mrch1 = 'DS'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '9000'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8050'; $dept_grp5 = '8060'; $dept_grp6 = '8070'; $dept_grp7 = '0000';
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'DS2001'; $store2 = 'DS2002'; $store3 = 'DS2003'; $store4 = 'DS2004'; $store5 = 'DS2006'; $store6 = 'DS2007'; $store7 = 'DS2009'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 

	elsif($type_test eq 2) {	
		
		&query_by_store;	
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
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'DS2223'; $store2 = '0000'; $store3 = '0000'; $store4 = '0000'; $store5 = '0000'; $store6 = '0000'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2223'; $stor2 = '0000'; $stor3 = '0000'; $stor4 = '0000'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

if($type_test eq 1){	&query_summary;	} 

elsif($type_test eq 2) {	
	
	&query_by_store;	
	&calc8;
	
	$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
	$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
	$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

	$new_ds=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

		foreach my $col( 7, 8, 10, 12, 13, 15 ){
			my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_ds,$col).','.xl_rowcol_to_cell($new_ds,$col).')';
				$worksheet->write( $a, $col, $sumTY, $bodyNum );
				
				if ($col eq 8 or $col eq 13){
					if( xl_rowcol_to_cell( $a, $col ) le 0){
						$worksheet->write( $a, $col+1, '0', $bodyPct );
					}
					else{
						my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
						$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
					}
				}
				
				# if ($col eq 10 or $col eq 15){
					# if( xl_rowcol_to_cell( $a, $col-3 ) le 0 or xl_rowcol_to_cell( $a, $col ) le 0){
						# $worksheet->write( $a, $col+1, '0', $bodyPct );
					# }
					# else{
						# my $pct2sls = '=IFERROR(('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1 ),"")';
						# $worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
					#}
				# }
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
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;

$store1 = 'DS2001'; $store2 = 'DS2002'; $store3 = 'DS2003'; $store4 = 'DS2004'; $store5 = 'DS2006'; $store6 = 'DS2007'; $store7 = 'DS2009'; $store8 = 'DS2223'; $store9 = '0000'; $store10 = 'OOOO'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2223'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

&query_summary;	

}

sub str_Su {

$div_name = "Comp"; $div_name3 = 'Supermarket';
$mrch1 = 'SU'; $mrch2 = 'SU'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8040'; $dept_grp2 = '0000'; $dept_grp3 = '0000'; $dept_grp4 = '0000'; $dept_grp5 = '0000'; $dept_grp6 = '0000'; $dept_grp7 = '0000';
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;

$store1 = 'SU2001'; $store2 = 'SU2002'; $store3 = 'SU2003'; $store4 = 'SU2004'; $store5 = '0000'; $store6 = 'SU2006'; $store7 = 'SU2007'; $store8 = '0000'; $store9 = 'SU2009'; $store10 = 'SU2013'; $store11 = 'SU4004'; $store12 = 'SU2012'; $store13 = 'SU3009'; $store14 = 'SU3010'; $store15 = 'SU3011'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2013'; $stor9 = '2012'; $stor10 = '4004'; $stor11 = '3009'; $stor12 = '3010';  $stor13 = '3011'; 

&query_summary;
	
}

sub str_Hy {

$div_name = "New"; $div_name3 = 'Hypermarket';
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU6001'; $store2 = 'SU6002'; $store3 = 'SU6003'; $store4 = 'SU6004'; $store5 = 'SU6005'; $store6 = 'SU6012'; $store7 = 'SU6009'; $store8 = 'SU6010'; $store9 = 'SU6011'; $store10 = 'DS6001'; $store11 = 'DS6002'; $store12 = 'DS6003'; $store13 = 'DS6004'; $store14 = 'DS6005'; $store15 = 'DS6012'; $store16 = 'DS6009'; $store17 = 'DS6010'; $store18 = 'DS6011'; $store19 = 'OT6002'; $store20 = 'OT6003'; $store21 = 'OT6004'; $store22 = 'OT6005'; $store23 = 'OT6012'; $store24 = 'OT6009'; $store25 = 'OT6010'; $store26 = 'OT6011';  $store27 = 'OT6000'; $store28 = 'SU2005'; $store29 = 'SU2008'; $store30 = 'SU2010'; $store31 = 'SU2011'; $store32 = 'DS2005'; $store33 = 'DS2008'; $store34 = 'DS2010'; $store35 = 'DS2011'; $store36 = 'DS2005'; $store37 = 'OT2005'; $store38 = 'OT2008'; $store39 = 'OT2010'; $store40 = 'OT2011';

$stor1 = '6001'; $stor2 = '6002'; $stor3 = '6003'; $stor4 = '6004'; $stor5 = '6005'; $stor6 = '6012'; $stor7 = '6009'; $stor8 = '6010'; $stor9 = '6011'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000'; 

&query_summary;	

}

sub str_Nb {

$div_name = "All"; $div_name3 = 'Neighborhood Store';
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000';  $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU3000'; $store2 = 'SU3007'; $store3 = 'SU4003'; $store4 = 'DS3000'; $store5 = 'DS3007'; $store6 = 'DS4003'; $store7 = 'OT3000'; $store8 = 'OT3007'; $store9 = 'OT4003'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '3000'; $stor2 = '3007'; $stor3 = '4003'; $stor4 = '0000'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';

&query_summary;	
$counter = 4; 
&calc8; 

$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
$worksheet->merge_range( $a-$counter, 3, $a, 3, $div_name, $border2 );

$a+=7; 
$counter = 0; 
$d=$a;

}


sub new_sheet{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(100);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
$worksheet->set_print_scale( 100 );
$worksheet->set_margins( 0.05 );
$worksheet->conditional_formatting( 'F9:V2000',  { type => 'cell', criteria => '<', value => 0, format => $down } );
$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
#$worksheet->set_column( 5, 5, 5 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 23 );
$worksheet->set_column( 7, 9, 8 );
$worksheet->set_column( 12, 14, 8 );
$worksheet->set_column( 10, 11, undef, undef, 1 );
$worksheet->set_column( 15, 16, undef, undef, 1 );

}

sub heading {

$worksheet->write($a-3, 3, "in 000's", $script);
$worksheet->merge_range( $a-2, 7, $a-2, 11, 'WTD', $subhead );
$worksheet->merge_range( $a-2, 12, $a-2, 16, 'MTD', $subhead );

foreach my $i ( 7, 12 ) {
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


sub query_dept {

$table = 'bi_sales_perf.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls1 = $dbh_csv->prepare (qq{SELECT merch_group_code_rev, SUM(target_sale_val) AS wtd_target, SUM(net_sale_ty) AS wtd_net_ty, SUM(net_sale_ly) AS wtd_net_ly, SUM(mtd_target_sale_val) AS mtd_target, SUM(mtd_net_sale_ty) AS mtd_net_ty, SUM(mtd_net_sale_ly) AS mtd_net_ly
							  FROM $table
							  WHERE new_flg = '$new_flg' or matured_flg = '$matured_flg'
							  GROUP BY merch_group_code_rev
							  ORDER BY merch_group_code_rev
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{merch_group_code_rev};
	#$merch_group_desc = $s->{merch_group_desc};
	
	$sls2 = $dbh_csv->prepare (qq{SELECT group_code, group_desc, SUM(target_sale_val) AS wtd_target, SUM(net_sale_ty) AS wtd_net_ty, SUM(net_sale_ly) AS wtd_net_ly, SUM(mtd_target_sale_val) AS mtd_target, SUM(mtd_net_sale_ty) AS mtd_net_ty, SUM(mtd_net_sale_ly) AS mtd_net_ly
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
				
		$sls3 = $dbh_csv->prepare (qq{SELECT division, division_desc, SUM(target_sale_val) AS wtd_target, SUM(net_sale_ty) AS wtd_net_ty, SUM(net_sale_ly) AS wtd_net_ly, SUM(mtd_target_sale_val) AS mtd_target, SUM(mtd_net_sale_ty) AS mtd_net_ty, SUM(mtd_net_sale_ly) AS mtd_net_ly
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
			
			$sls4 = $dbh_csv->prepare (qq{SELECT department_code, department_desc, SUM(target_sale_val) AS wtd_target, SUM(net_sale_ty) AS wtd_net_ty, SUM(net_sale_ly) AS wtd_net_ly, SUM(mtd_target_sale_val) AS mtd_target, SUM(mtd_net_sale_ty) AS mtd_net_ty, SUM(mtd_net_sale_ly) AS mtd_net_ly 
										 FROM $table 
										 WHERE group_code = '$group_code' and merch_group_code_rev = '$merch_group_code' and division = '$division' and (new_flg = '$new_flg' or matured_flg = '$matured_flg')
										 GROUP BY department_code, department_desc 
										 ORDER BY department_code
										 });
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){
				
				$worksheet->write($a,5, $s->{department_code},$desc);
				$worksheet->write($a,6, $s->{department_desc},$desc);
				$worksheet->write($a,7, $s->{wtd_net_ty},$border1);
				$worksheet->write($a,8, $s->{wtd_net_ly},$border1);
				$worksheet->write($a,10, $s->{wtd_target},$border1);
				if ($s->{wtd_net_ly} <= 0){
					$worksheet->write($a,9, "",$subt); }
				else{
					$worksheet->write($a,9, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); }
				
				if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0 ){
					$worksheet->write($a,11, "",$subt); }
				else{
					$worksheet->write($a,11, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); }
					
				$worksheet->write($a,12, $s->{mtd_net_ty},$border1);
				$worksheet->write($a,13, $s->{mtd_net_ly},$border1);
				$worksheet->write($a,15, $s->{mtd_target},$border1);
				if ($s->{mtd_net_ly} <= 0){
					$worksheet->write($a,14, "",$subt); }
				else{
					$worksheet->write($a,14, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); }
				
				if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0 ){
					$worksheet->write($a,16, "",$subt); }
				else{
					$worksheet->write($a,16, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); }
				
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
			$grp_wtd_net_ty += $s->{wtd_net_ty};
			$grp_wtd_net_ly += $s->{wtd_net_ly};
			$grp_wtd_target += $s->{wtd_target};
			
			$grp_mtd_net_ty += $s->{mtd_net_ty};
			$grp_mtd_net_ly += $s->{mtd_net_ly};
			$grp_mtd_target += $s->{mtd_target};
		}
		
		$worksheet->write($a,7, $s->{wtd_net_ty},$bodyNum);
		$worksheet->write($a,8, $s->{wtd_net_ly},$bodyNum);
		$worksheet->write($a,10, $s->{wtd_target},$bodyNum);
		if ($s->{wtd_net_ly} <= 0){
			$worksheet->write($a,9, "",$bodyPct); }
		else{
			$worksheet->write($a,9, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$bodyPct); }

		if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0 ){
			$worksheet->write($a,11, "",$bodyPct); }
		else{
			$worksheet->write($a,11, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$bodyPct); }
			
		$worksheet->write($a,12, $s->{mtd_net_ty},$bodyNum);
		$worksheet->write($a,13, $s->{mtd_net_ly},$bodyNum);
		$worksheet->write($a,15, $s->{mtd_target},$bodyNum);
		if ($s->{mtd_net_ly} <= 0){
			$worksheet->write($a,14, "",$bodyPct); }
		else{
			$worksheet->write($a,14, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$bodyPct); }
			
		if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0 ){
			$worksheet->write($a,16, "",$bodyPct); }
		else{
			$worksheet->write($a,16, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$bodyPct); }

		$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

		$a++; #INCREMENT VARIABLE a
	}
	
	$total_wtd_net_ty += $s->{wtd_net_ty};
	$total_wtd_net_ly += $s->{wtd_net_ly};
	$total_wtd_target += $s->{wtd_target};
	
	$total_mtd_net_ty += $s->{mtd_net_ty};
	$total_mtd_net_ly += $s->{mtd_net_ly};
	$total_mtd_target += $s->{mtd_target};
	
	if ($merch_group_code eq 'DS'){
		############DEPT STORE WO JEWELRY ###############
		$worksheet->write($a,7, $grp_wtd_net_ty,$headNum);
		$worksheet->write($a,8, $grp_wtd_net_ly,$headNum);
		$worksheet->write($a,10, $grp_wtd_target,$headNum);
			if ($grp_wtd_net_ly <= 0){
				$worksheet->write($a,9, "",$headDPct); }
			else{
				$worksheet->write($a,9, ($grp_wtd_net_ty-$grp_wtd_net_ly)/$grp_wtd_net_ly,$headDPct); }
				
			if ($grp_wtd_net_ty <= 0 or $grp_wtd_target <= 0 ){
				$worksheet->write($a,11, "",$headDPct); }
			else{
				$worksheet->write($a,11, ($grp_wtd_net_ty-$grp_wtd_target)/$grp_wtd_target,$headDPct); }
		
		$worksheet->write($a,12, $grp_mtd_net_ty,$headNum);
		$worksheet->write($a,13, $grp_mtd_net_ly,$headNum);
		$worksheet->write($a,15, $grp_mtd_target,$headNum);
			if ($grp_mtd_net_ly <= 0){
				$worksheet->write($a,14, "",$headDPct); }
			else{
				$worksheet->write($a,14, ($grp_mtd_net_ty-$grp_mtd_net_ly)/$grp_mtd_net_ly,$headDPct); }
				
			if ($grp_mtd_net_ty <= 0 or $grp_mtd_target <= 0 ){
				$worksheet->write($a,16, "",$headPct); }
			else{
				$worksheet->write($a,16, ($grp_mtd_net_ty-$grp_mtd_target)/$grp_mtd_target,$headDPct); }
		
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store w/o Jewelry', $headD );
	
		$a += 1;
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );}
	elsif($merch_group_code eq 'SU'){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );}
	
	$worksheet->write($a,7, $s->{wtd_net_ty},$headNumber);
	$worksheet->write($a,8, $s->{wtd_net_ly},$headNumber);
	$worksheet->write($a,10, $s->{wtd_target},$headNumber);
		if ($s->{wtd_net_ly} <= 0){
			$worksheet->write($a,9, "",$headPct); }
		else{
			$worksheet->write($a,9, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$headPct); }
			
		if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0 ){
			$worksheet->write($a,11, "",$headPct); }
		else{
			$worksheet->write($a,11, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$headPct); }
	
	$worksheet->write($a,12, $s->{mtd_net_ty},$headNumber);
	$worksheet->write($a,13, $s->{mtd_net_ly},$headNumber);
	$worksheet->write($a,15, $s->{mtd_target},$headNumber);
		if ($s->{mtd_net_ly} <= 0){
			$worksheet->write($a,14, "",$headPct); }
		else{
			$worksheet->write($a,14, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$headPct); }
			
		if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0 ){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$headPct); }
	
	$a++; #INCREMENT VARIABLE a
}

	$worksheet->write($a,7, $total_wtd_net_ty,$headNumber);
	$worksheet->write($a,8, $total_wtd_net_ly,$headNumber);
	$worksheet->write($a,10, $total_wtd_target,$headNumber);
		if ($total_wtd_net_ly <= 0){
			$worksheet->write($a,9, "",$headPct); }
		else{
			$worksheet->write($a,9, ($total_wtd_net_ty-$total_wtd_net_ly)/$total_wtd_net_ly,$headPct); }
			
		if ($total_wtd_net_ty <= 0 or $total_wtd_target <= 0 ){
			$worksheet->write($a,11, "",$headPct); }
		else{
			$worksheet->write($a,11, ($total_wtd_net_ty-$total_wtd_target)/$total_wtd_target,$headPct); }
	
	$worksheet->write($a,12, $total_mtd_net_ty,$headNumber);
	$worksheet->write($a,13, $total_mtd_net_ly,$headNumber);
	$worksheet->write($a,15, $total_mtd_target,$headNumber);
		if ($total_mtd_net_ly <= 0){
			$worksheet->write($a,14, "",$headPct); }
		else{
			$worksheet->write($a,14, ($total_mtd_net_ty-$total_mtd_net_ly)/$total_mtd_net_ly,$headPct); }
			
		if ($total_mtd_net_ty <= 0 or $total_mtd_target <= 0 ){
			$worksheet->write($a,16, "",$headPct); }
		else{
			$worksheet->write($a,16, ($total_mtd_net_ty-$total_mtd_target)/$total_mtd_target,$headPct); }
	
$worksheet->write($loc, 2, $loc_desc, $bold);
$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

}


sub query_summary {

$table = 'bi_sales_perf.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls = $dbh_csv->prepare (qq{SELECT SUM(target_sale_val) AS wtd_target, SUM(net_sale_ty) AS wtd_net_ty, SUM(net_sale_ly) AS wtd_net_ly, SUM(mtd_target_sale_val) AS mtd_target, SUM(mtd_net_sale_ty) AS mtd_net_ty, SUM(mtd_net_sale_ly) AS mtd_net_ly
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (department_code = '$dept_grp1' or department_code = '$dept_grp2' 
																										or department_code = '$dept_grp3' or department_code = '$dept_grp4' 
																										or department_code = '$dept_grp5' or department_code = '$dept_grp6' 
																										or department_code = '$dept_grp7')) or
																		division = '$division_grp3'))
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13'))
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
	$worksheet->write($a,7, $s->{wtd_net_ty},$border1);
	$worksheet->write($a,8, $s->{wtd_net_ly},$border1);
	$worksheet->write($a,10, $s->{wtd_target},$border1);
	if ($s->{wtd_net_ly} <= 0){
		$worksheet->write($a,9, "",$subt);
	}
	else{
		$worksheet->write($a,9, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt);
	}
	if ($s->{net_ty} <= 0 or $s->{wtd_target} <= 0){
		$worksheet->write($a,11, "",$subt);
	}
	else{
		$worksheet->write($a,11, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt);
	}
	
	$worksheet->write($a,12, $s->{mtd_net_ty},$border1);
	$worksheet->write($a,13, $s->{mtd_net_ly},$border1);
	$worksheet->write($a,15, $s->{mtd_target},$border1);
	if ($s->{mtd_net_ly} <= 0){
		$worksheet->write($a,14, "",$subt);
	}
	else{
		$worksheet->write($a,14, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt);
	}
	if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0){
		$worksheet->write($a,16, "",$subt);
	}
	else{
		$worksheet->write($a,16, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt);
	}
	
	$a++;
	$counter++;
}

$sls->finish();

}

sub query_by_store {

$table = 'bi_sales_perf.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;


$sls = $dbh_csv->prepare (qq{SELECT store_code, store_description, SUM(target_sale_val) AS wtd_target, SUM(net_sale_ty) AS wtd_net_ty, SUM(net_sale_ly) AS wtd_net_ly, SUM(mtd_target_sale_val) AS mtd_target, SUM(mtd_net_sale_ty) AS mtd_net_ty, SUM(mtd_net_sale_ly) AS mtd_net_ly
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (department_code = '$dept_grp1' or department_code = '$dept_grp2' 
																										or department_code = '$dept_grp3' or department_code = '$dept_grp4' 
																										or department_code = '$dept_grp5' or department_code = '$dept_grp6' 
																										or department_code = '$dept_grp7')) or
																		division = '$division_grp3')) 
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13'))
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
	if ($mrch1 eq 'SU' and $mrch2 eq 'SU' and $s->{store_code} eq '2001') {$s->{store_description} = 'METRO COLON + F2';}
	$worksheet->write($a,5, $s->{store_code},$desc);
	$worksheet->write($a,6, $s->{store_description},$desc);
	$worksheet->write($a,7, $s->{wtd_net_ty},$border1);
	$worksheet->write($a,8, $s->{wtd_net_ly},$border1);
	$worksheet->write($a,10, $s->{wtd_target},$border1);
	if ($s->{wtd_net_ly} <= 0){
		$worksheet->write($a,9, "",$subt); }
	else{
		$worksheet->write($a,9, ($s->{wtd_net_ty}-$s->{wtd_net_ly})/$s->{wtd_net_ly},$subt); }
	
	if ($s->{wtd_net_ty} <= 0 or $s->{wtd_target} <= 0 ){
		$worksheet->write($a,11, "",$subt); }
	else{
		$worksheet->write($a,11, ($s->{wtd_net_ty}-$s->{wtd_target})/$s->{wtd_target},$subt); }
		
	$worksheet->write($a,12, $s->{mtd_net_ty},$border1);
	$worksheet->write($a,13, $s->{mtd_net_ly},$border1);
	$worksheet->write($a,15, $s->{mtd_target},$border1);
	if ($s->{mtd_net_ly} <= 0){
		$worksheet->write($a,14, "",$subt); }
	else{
		$worksheet->write($a,14, ($s->{mtd_net_ty}-$s->{mtd_net_ly})/$s->{mtd_net_ly},$subt); }
	
	if ($s->{mtd_net_ty} <= 0 or $s->{mtd_target} <= 0 ){
		$worksheet->write($a,16, "",$subt); }
	else{
		$worksheet->write($a,16, ($s->{mtd_net_ty}-$s->{mtd_target})/$s->{mtd_target},$subt); }
		
	$a++;
	$counter++;
}

$sls->finish();

}


sub calc8 { 

foreach my $col( 7, 8, 10, 12, 13, 15 ){
	my $sum = '=SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . ')';
		$worksheet->write( $a, $col, $sum, $bodyNum );
		
		if ($col eq 8 or $col eq 13){
			if( xl_rowcol_to_cell( $a, $col ) le 0){
				$worksheet->write( $a, $col+1, '0', $bodyPct );
			}
			else{
				my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
				$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
			}
		}
		
		#if ($col eq 10 or $col eq 15){
			# if( xl_rowcol_to_cell( $a, $col-3 ) le 0 or xl_rowcol_to_cell( $a, $col ) le 0){
				# $worksheet->write( $a, $col+1, '0', $bodyPct );
			# }
			# else{
				#my $pct2sls = '=IFERROR(('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1), "")';
				#$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
			# }
		#}
}

}

sub calc9 { 

foreach my $col( 7, 8, 10, 12, 13, 15 ){
	$sumTY = '=SUM('.xl_rowcol_to_cell($comp_su,$col).','.xl_rowcol_to_cell($new_su,$col).','.xl_rowcol_to_cell($comp_nb,$col).','.xl_rowcol_to_cell($comp_hy,$col).','.xl_rowcol_to_cell($new_hy,$col).','.xl_rowcol_to_cell($comp_ds,$col).','.xl_rowcol_to_cell($new_ds,$col).')';
		$worksheet->write( $a, $col, $sumTY, $headNumber );
		
		if ($col eq 8 or $col eq 13){
			if( xl_rowcol_to_cell( $a, $col ) le 0){
				$worksheet->write( $a, $col+1, '0', $head );
			}
			else{
				my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
				$worksheet->write( $a, $col+1, $pct2sls, $headPct );
			}
		}
		
		if ($col eq 10 or $col eq 15){
			if( xl_rowcol_to_cell( $a, $col-3 ) le 0){
				$worksheet->write( $a, $col+1, '0', $head );
			}
			else{
				my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-3 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
				$worksheet->write( $a, $col+1, $pct2sls, $headPct );
			}
		}
}

$worksheet->merge_range( $a, 3, $a, 6, "TOTAL", $headN );

$a+=1; $counter = 0; $d=$a;

}


sub generate_csv {

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh, ">", "bi_sales_perf.csv" or die "bi_sales_perf.csv: $!";

$test = qq{ 
SELECT 
BASE.STORE_FORMAT, 
BASE.STORE_FORMAT_DESC, 
BASE.STORE_CODE, 
CASE WHEN BASE.STORE_CODE IN (2012, 2013, 3009, 4004, 3010, 3011) THEN 'SU' || BASE.STORE_CODE 
     WHEN BASE.STORE_CODE = 2223 THEN 'DS' || BASE.STORE_CODE 
	 ELSE BASE.MERCH_GROUP_CODE || BASE.STORE_CODE END AS STORE,	 
UPPER(BASE.STORE_DESCRIPTION) AS STORE_DESCRIPTION,
CASE WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 9000) THEN 'DS'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8500) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE = 8040) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE != 8040) THEN 'DS'
ELSE BASE.MERCH_GROUP_CODE END AS MERCH_GROUP_CODE_REV,
BASE.MERCH_GROUP_CODE, 
BASE.MERCH_GROUP_DESC, 
BASE.GROUP_CODE, 
BASE.GROUP_DESC, 
BASE.DIVISION, 
BASE.DIVISION_DESC, 
BASE.DEPARTMENT_CODE, 
BASE.DEPARTMENT_DESC, 
BASE.NEW_FLG, 
BASE.MATURED_FLG,
SUM(WTD.TARGET_SALE_VAL) TARGET_SALE_VAL, SUM(WTD.NET_SALE_TY) NET_SALE_TY, SUM(WTD.NET_SALE_LY) NET_SALE_LY
, SUM(MTD.TARGET_SALE_VAL) MTD_TARGET_SALE_VAL, SUM(MTD.NET_SALE_TY) MTD_NET_SALE_TY, SUM(MTD.NET_SALE_LY) MTD_NET_SALE_LY
FROM
(SELECT S.STORE_FORMAT, S.STORE_FORMAT_DESC, 
	CASE WHEN S.STORE_CODE IN ('4002') THEN '2001' 
		WHEN S.STORE_CODE IN ('3001','3002','3003','3004','3005','3006') THEN '3000'
		ELSE S.STORE_CODE END AS STORE_CODE, 
	CASE WHEN (S.STORE_DESCRIPTION LIKE 'Metro Wholesalemart Colon%') THEN 'METRO COLON' 
		WHEN (S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Punta%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Basak%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Minglanilla%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Tabunok%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Tabok%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Umapad%') THEN 'METRO FRESH N EASY'
		ELSE UPPER(S.STORE_DESCRIPTION) END AS STORE_DESCRIPTION, 
	M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.GROUP_CODE, M.GROUP_DESC, M.DIVISION, M.DIVISION_DESC, M.DEPARTMENT_CODE, M.DEPARTMENT_DESC,S.NEW_FLG, S.MATURED_FLG
FROM
	(SELECT STORE_FORMAT, STORE_FORMAT_DESC, STORE_CODE, STORE_DESCRIPTION, NEW_FLG, MATURED_FLG FROM DIM_STORE WHERE ACTIVE = 1 AND STORE_FORMAT IN (1, 2, 3, 4, 5) and STORE_CODE NOT IN (6008))S,
	(SELECT D.MERCH_GROUP_CODE, D.MERCH_GROUP_DESC, D.GROUP_CODE, D.GROUP_DESC, D.DIVISION, D.DIVISION_DESC, D.DEPARTMENT_CODE, D.DEPARTMENT_DESC 
		FROM DIM_MERCHANDISE M 
			JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE 
								AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE )M
GROUP BY S.STORE_FORMAT, S.STORE_FORMAT_DESC, 
	CASE WHEN S.STORE_CODE IN ('4002') THEN '2001' 
		WHEN S.STORE_CODE IN ('3001','3002','3003','3004','3005','3006') THEN '3000'
		ELSE S.STORE_CODE END, 
	CASE WHEN (S.STORE_DESCRIPTION LIKE 'Metro Wholesalemart Colon%') THEN 'METRO COLON'
		WHEN (S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Punta%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Basak%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Minglanilla%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Tabunok%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Tabok%' OR S.STORE_DESCRIPTION LIKE 'Metro Fresh n Easy Umapad%') THEN 'METRO FRESH N EASY'
		ELSE UPPER(S.STORE_DESCRIPTION) END, 
	M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.GROUP_CODE, M.GROUP_DESC, M.DIVISION, M.DIVISION_DESC, M.DEPARTMENT_CODE, M.DEPARTMENT_DESC,S.NEW_FLG, S.MATURED_FLG)BASE
LEFT JOIN
(SELECT 
DIM_STORE.STORE_FORMAT, 
CASE WHEN  DIM_STORE.STORE_CODE IN ('4002') THEN '2001' 
	WHEN DIM_STORE.STORE_CODE IN ('3001','3002','3003','3004','3005','3006') THEN '3000'
	ELSE DIM_STORE.STORE_CODE END AS STORE_CODE,
DIM_SUB_DEPT.MERCH_GROUP_CODE,
DIM_SUB_DEPT.GROUP_CODE,
DIM_SUB_DEPT.DIVISION,
DIM_SUB_DEPT.DEPARTMENT_CODE,
NVL((SUM(TARGET_SALE_VAL))/1000,0) TARGET_SALE_VAL, 
NVL((SUM(SALE_NET_VAL))/1000,0) SALE_NET_VAL, 
NVL((SUM(SALE_TOT_TAX_VAL))/1000,0) SALE_TOT_TAX_VAL,
NVL((SUM(SALE_TOT_DISC_VAL))/1000,0) SALE_TOT_DISC_VAL, 
NVL((SUM(SALE_TOT_DISC_VAL_LY))/1000,0) SALE_TOT_DISC_VAL_LY, 
NVL((SUM((NVL(SALE_NET_VAL,0))-(NVL(SALE_TOT_TAX_VAL,0))-(NVL(SALE_TOT_DISC_VAL,0))))/1000,0) NET_SALE_TY,
NVL((SUM ((NVL(SALE_NET_VAL_LY,0))-(NVL(SALE_TOT_TAX_VAL_LY,0))-(NVL(SALE_TOT_DISC_VAL_LY,0))))/1000,0) NET_SALE_LY
FROM 
	(SELECT
	DATE_KEY, STORE_KEY, DS_KEY, STORE_CODE, 
	SUM (SALE_NET_VAL) AS SALE_NET_VAL, 
	SUM (SALE_NET_VAL_LY) AS SALE_NET_VAL_LY, 
	SUM (SALE_TOT_TAX_VAL) AS SALE_TOT_TAX_VAL, 
	SUM (SALE_TOT_TAX_VAL_LY) AS SALE_TOT_TAX_VAL_LY, 
	SUM (SALE_TOT_DISC_VAL) AS SALE_TOT_DISC_VAL,
	SUM (SALE_TOT_DISC_VAL_LY) AS SALE_TOT_DISC_VAL_LY,
	SUM (TARGET_SALE_VAL) AS TARGET_SALE_VAL, 
	SUM (TARGET_SALE_VAL_LY) AS TARGET_SALE_VAL_LY, 
	SUM (TARGET_SALE_VAT) AS TARGET_SALE_VAT, 
	SUM (TARGET_SALE_VAT_LY) AS TARGET_SALE_VAT_LY 
	FROM (SELECT
		DATE_KEY, STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE, 
		SUM(SALE_NET_VAL) AS SALE_NET_VAL, 
		SUM(SALE_NET_VAL_LY) AS SALE_NET_VAL_LY, 
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL, 
		SUM(SALE_TOT_TAX_VAL_LY) SALE_TOT_TAX_VAL_LY, 
		SUM(SALE_TOT_DISC_VAL) SALE_TOT_DISC_VAL,
		SUM(SALE_TOT_DISC_VAL_LY) SALE_TOT_DISC_VAL_LY, 0 AS TARGET_SALE_VAL, 0 AS TARGET_SALE_VAL_LY, 0 AS TARGET_SALE_VAT, 0 AS TARGET_SALE_VAT_LY 
		FROM AGG_WLY_STR_DCS AGG JOIN DIM_MERCHANDISE M ON AGG.DCS_KEY = M.DCS_KEY JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE  
		WHERE DATE_KEY BETWEEN $wk_st_date_key AND $wk_en_date_key 
		GROUP BY DATE_KEY, STORE_KEY, DS_KEY, STORE_CODE 
		UNION ALL 
		SELECT
		DATE_KEY, STORE_KEY, DS_KEY, STORE_CODE, 0 AS SALE_NET_VAL, 0 AS SALE_NET_VAL_LY, 0 AS SALE_TOT_TAX_VAL, 0 AS SALE_TOT_TAX_VAL_LY, 0 AS SALE_TOT_DISC_VAL, 0 AS SALE_TOT_DISC_VAL_LY,
		SUM (TARGET_SALE_VAL) AS TARGET_SALE_VAL, 
		SUM (TARGET_SALE_VAL_LY) AS TARGET_SALE_VAL_LY, 
		SUM (TARGET_SALE_VAT) AS TARGET_SALE_VAT, 
		SUM (TARGET_SALE_VAT_LY) AS TARGET_SALE_VAT_LY 
		FROM FCT_TARGET A 
		WHERE A.DATE_KEY BETWEEN $wk_st_date_key AND $wk_en_date_key 
		GROUP BY DATE_KEY, STORE_KEY, STORE_CODE, DS_KEY ) FINAL 
	GROUP BY DATE_KEY, STORE_KEY, DS_KEY, STORE_CODE) AGG_DLY_STR_DEPT_TARGET,DIM_STORE,DIM_SUB_DEPT 
WHERE AGG_DLY_STR_DEPT_TARGET.DATE_KEY >=$wk_st_date_key AND AGG_DLY_STR_DEPT_TARGET.DATE_KEY <=$wk_en_date_key AND DIM_STORE.ACTIVE = 1 AND DIM_STORE.STORE_FORMAT IN (1, 2, 3, 4, 5) AND AGG_DLY_STR_DEPT_TARGET.STORE_KEY=DIM_STORE.STORE_KEY AND AGG_DLY_STR_DEPT_TARGET.DS_KEY=DIM_SUB_DEPT.DS_KEY 
GROUP BY DIM_STORE.STORE_FORMAT, 
	CASE WHEN  DIM_STORE.STORE_CODE IN ('4002') THEN '2001' 
		WHEN DIM_STORE.STORE_CODE IN ('3001','3002','3003','3004','3005','3006') THEN '3000'
		ELSE DIM_STORE.STORE_CODE END, 
	DIM_SUB_DEPT.MERCH_GROUP_CODE, DIM_SUB_DEPT.GROUP_CODE, DIM_SUB_DEPT.DIVISION, DIM_SUB_DEPT.DEPARTMENT_CODE
)WTD
ON BASE.STORE_FORMAT = WTD.STORE_FORMAT AND BASE.STORE_CODE = WTD.STORE_CODE AND BASE.MERCH_GROUP_CODE = WTD.MERCH_GROUP_CODE AND BASE.GROUP_CODE = WTD.GROUP_CODE AND BASE.DIVISION = WTD.DIVISION AND BASE.DEPARTMENT_CODE = WTD.DEPARTMENT_CODE
LEFT JOIN
(SELECT
DIM_STORE.STORE_FORMAT, 
CASE WHEN DIM_STORE.STORE_CODE IN ('4002') THEN '2001' 
	WHEN DIM_STORE.STORE_CODE IN ('3001','3002','3003','3004','3005','3006') THEN '3000'
	ELSE DIM_STORE.STORE_CODE END AS STORE_CODE,
DIM_SUB_DEPT.MERCH_GROUP_CODE,
DIM_SUB_DEPT.GROUP_CODE,
DIM_SUB_DEPT.DIVISION,
DIM_SUB_DEPT.DEPARTMENT_CODE,
0 AS TARGET_SALE_VAL, 
NVL((SUM(SALE_NET_VAL))/1000,0) SALE_NET_VAL, 
NVL((SUM(SALE_TOT_TAX_VAL))/1000,0) SALE_TOT_TAX_VAL,
NVL((SUM(SALE_TOT_DISC_VAL))/1000,0) SALE_TOT_DISC_VAL, 
NVL((SUM(SALE_TOT_DISC_VAL_LY))/1000,0) SALE_TOT_DISC_VAL_LY, 
NVL((SUM((NVL(SALE_NET_VAL,0))-(NVL(SALE_TOT_TAX_VAL,0))-(NVL(SALE_TOT_DISC_VAL,0))))/1000,0) NET_SALE_TY,
NVL((SUM ((NVL(SALE_NET_VAL_LY,0))-(NVL(SALE_TOT_TAX_VAL_LY,0))-(NVL(SALE_TOT_DISC_VAL_LY,0))))/1000,0) NET_SALE_LY
FROM (
	SELECT TBL.STORE_KEY, TBL.STORE_CODE, TBL.DS_KEY, TY.SALE_NET_VAL, TY.SALE_TOT_TAX_VAL, TY.SALE_TOT_DISC_VAL, LY.SALE_NET_VAL_LY, LY.SALE_TOT_TAX_VAL_LY, LY.SALE_TOT_DISC_VAL_LY FROM
		(SELECT S.STORE_KEY, S.STORE_CODE, M.DS_KEY
		FROM
			(SELECT STORE_KEY, STORE_CODE
			FROM DIM_STORE 
			WHERE ACTIVE = 1 AND STORE_FORMAT IN (1, 2, 3, 4, 5))S,
			(SELECT D.DS_KEY, D.MERCH_GROUP_CODE, D.MERCH_GROUP_DESC, D.GROUP_CODE, D.GROUP_DESC, D.DIVISION, D.DIVISION_DESC, D.DEPARTMENT_CODE, D.DEPARTMENT_DESC
			FROM DIM_MERCHANDISE M 
				JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION 
					AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE )M
		GROUP BY S.STORE_KEY, S.STORE_CODE, M.DS_KEY)TBL
		LEFT JOIN
		(SELECT
		STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE,  
		SUM(SALE_NET_VAL) AS SALE_NET_VAL, 
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL, 
		SUM(SALE_TOT_DISC_VAL) SALE_TOT_DISC_VAL
		FROM AGG_DLY_STR_PROD AGG 
			JOIN DIM_PRODUCT M ON AGG.PRODUCT_KEY = M.PRODUCT_KEY 
			JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
		WHERE DATE_KEY BETWEEN $mo_st_date_key AND $mo_en_date_key 
		GROUP BY STORE_KEY, DS_KEY, STORE_CODE)TY
		ON TBL.STORE_KEY = TY.STORE_KEY AND TBL.STORE_CODE = TY.STORE_CODE AND TBL.DS_KEY = TY.DS_KEY
		LEFT JOIN
		(SELECT
		STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE,  
		SUM(SALE_NET_VAL) AS SALE_NET_VAL_LY, 
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_LY, 
		SUM(SALE_TOT_DISC_VAL) SALE_TOT_DISC_VAL_LY
		FROM AGG_DLY_STR_PROD AGG 
			JOIN DIM_PRODUCT M ON AGG.PRODUCT_KEY = M.PRODUCT_KEY 
			JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
		WHERE DATE_KEY BETWEEN $mo_st_date_key_ly AND $mo_en_date_key_ly 
		GROUP BY STORE_KEY, DS_KEY, STORE_CODE)LY 
		ON TBL.STORE_KEY = LY.STORE_KEY AND TBL.STORE_CODE = LY.STORE_CODE AND TBL.DS_KEY = LY.DS_KEY ) AGG_MLY_STR_DEPT_TARGET,DIM_STORE,DIM_SUB_DEPT 
WHERE DIM_STORE.ACTIVE = 1 AND DIM_STORE.STORE_FORMAT IN (1, 2, 3, 4, 5) 
	AND AGG_MLY_STR_DEPT_TARGET.STORE_KEY=DIM_STORE.STORE_KEY AND AGG_MLY_STR_DEPT_TARGET.DS_KEY=DIM_SUB_DEPT.DS_KEY 
GROUP BY DIM_STORE.STORE_FORMAT, 
	CASE WHEN DIM_STORE.STORE_CODE IN ('4002') THEN '2001' 
		WHEN DIM_STORE.STORE_CODE IN ('3001','3002','3003','3004','3005','3006') THEN '3000'
		ELSE DIM_STORE.STORE_CODE END, 
	DIM_SUB_DEPT.MERCH_GROUP_CODE, DIM_SUB_DEPT.GROUP_CODE, DIM_SUB_DEPT.DIVISION, DIM_SUB_DEPT.DEPARTMENT_CODE
)MTD
ON BASE.STORE_FORMAT = MTD.STORE_FORMAT AND BASE.STORE_CODE = MTD.STORE_CODE AND BASE.MERCH_GROUP_CODE = MTD.MERCH_GROUP_CODE AND BASE.GROUP_CODE = MTD.GROUP_CODE AND BASE.DIVISION = MTD.DIVISION AND BASE.DEPARTMENT_CODE = MTD.DEPARTMENT_CODE
GROUP BY
BASE.STORE_FORMAT, 
BASE.STORE_FORMAT_DESC, 
BASE.STORE_CODE, 
CASE WHEN BASE.STORE_CODE IN (2012, 2013, 3009, 4004, 3010, 3011) THEN 'SU' || BASE.STORE_CODE
     WHEN BASE.STORE_CODE = 2223 THEN 'DS' || BASE.STORE_CODE 
	 ELSE BASE.MERCH_GROUP_CODE || BASE.STORE_CODE END,	 
UPPER(BASE.STORE_DESCRIPTION),
CASE WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 9000) THEN 'DS'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8500) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE = 8040) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE != 8040) THEN 'DS'
ELSE BASE.MERCH_GROUP_CODE END,
BASE.MERCH_GROUP_CODE, 
BASE.MERCH_GROUP_DESC, 
BASE.GROUP_CODE, 
BASE.GROUP_DESC, 
BASE.DIVISION, 
BASE.DIVISION_DESC, 
BASE.DEPARTMENT_CODE, 
BASE.DEPARTMENT_DESC, 
BASE.NEW_FLG, 
BASE.MATURED_FLG
};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_lc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "bi_sales_perf.csv: $!";
 
$sth->finish();
$dbh->disconnect;

}


sub mail_grp1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1, $attachment_file_2 ) = @ARGV;

$to = ' gerry.guanlao@metrogaisano.com, eric.redona@metrogaisano.com, lucille.malazarte@metrogaisano.com, tricia.luntao@metrogaisano.com, jj.moreno@metrogaisano.com, cj.jesena@metrogaisano.com ';

$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, christopher.calalang@metrogaisano.com, annalyn.conde@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com ';
		
# $to = 'fnaquines@metro.com.ph';
# $bcc = 'kent.mamalias@metrogaisano.com';
		
$subject = 'Daily Sales Performance as of ' . $as_of;
$msgbody_file = 'message_BI.txt';

$attachment_file_1 = "Daily Sales Performance - Summary (as of $as_of) v1.xlsx";
$attachment_file_2 = "Daily Sales Performance - Summary (as of $as_of) v1.pdf";

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

$to = ' cj.jesena@metrogaisano.com ';
$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, christopher.calalang@metrogaisano.com, annalyn.conde@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com ';
		
# $to = 'fnaquines@metro.com.ph';
# $bcc = 'kent.mamalias@metrogaisano.com';
		
$subject = 'Daily Sales Performance as of ' . $as_of;
$msgbody_file = 'message_BI.txt';

$attachment_file_1 = "Daily Sales Performance (as of $as_of) v1.xlsx";
$attachment_file_2 = "Daily Sales Performance (as of $as_of) v1.pdf";

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








