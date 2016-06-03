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
 
	$workbook = Excel::Writer::XLSX->new("New Store Daily Sales Performance - Summary (as of $as_of) v1.xlsx");
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
	$subhead = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, bg_color => $ponkan, bold => 1 );
	$bodyN = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10, bg_color => $sky, num_format => '0.0 %',  bold => 1);
	$bodyPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => '0.0 %',  bold => 1);
	$bodyNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 3,  bold => 1);
	$body = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => '0.0 %',  bold => 1);
	$subt = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => '0.0 %');
	$down = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => '0.0 %', bg_color => $pula );

	printf "Test ETL Status = ". $test ." \nArc BI Sales Performance Part 1\nGenerating Data from Source \n";
	
	$ns_st_date_key = 807;
	$ns_end_date_key = 836;
	
	&generate_csv;
	
	&new_sheet_dly($sheet = "Daily Trend");			
	&call_dly;
	
	&new_sheet_div($sheet = "Department");			
	&call_div;
	
	&new_sheet_div($sheet = "Merchandise");			
	&call_merch;
		
	$workbook->close();
	
	my $table = 'late_posted_'.$day . '_' . $month_to_text . '_' .$year . '.csv';

	my $dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
		or die $DBI::errstr;
	 
	my $sth = $dbh_csv->prepare(qq{SELECT SUM(VALUE) AS AMOUNT FROM $table WHERE STORE = '3010'});

	$sth->execute() or die "Failed to execute query - " . $dbh_csv->errstr;
	
	while(my $s = $sth->fetchrow_hashref()){
		 if ($s->{AMOUNT} ge 0){
			 &mail_grp1_lateposted; 
			 &mail_external_lateposted;
		}
		else{
			&mail_grp1;	
			&mail_external;
		}
	}
		
	$sth->finish();	
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

$a = 10, $counter = 0;
$total_net = 0;
$grp_net = 0; 

$worksheet->write($a-10, 2, "Daily Sales Performance", $bold1);
$worksheet->write($a-9, 2, "Metro Banilad", $bold);
$worksheet->write($a-8, 2, "As of $as_of");
$worksheet->write($a-3, 2, "in 000's", $script);

&heading_div;
&query_dept;

}

sub call_dly {

$a = 10, $counter = 0;

$worksheet->write($a-10, 5, "Daily Sales Performance", $bold1);
$worksheet->write($a-9, 5, "Metro Banilad", $bold);
$worksheet->write($a-8, 5, "As of $as_of");
$worksheet->write($a-3, 5, "in 000's", $script);

&heading_dly;
&query_dly;

}

sub call_merch {

$a = 10, $counter = 0;
$total_net = 0;
$grp_net = 0; 
$day_counter -= 1;

$worksheet->write($a-10, 2, "Daily Sales Performance", $bold1);
$worksheet->write($a-9, 2, "Metro Banilad", $bold);
$worksheet->write($a-8, 2, "As of $as_of");
$worksheet->write($a-3, 2, "in 000's", $script);

$worksheet->merge_range( $a-4, 7, $a-4, 12, 'DAY 1 TO ' . $day_counter, $subhead ); #
$worksheet->merge_range( $a-3, 7, $a-3, 9, 'SALES', $subhead );

$worksheet->merge_range( $a-2, 7, $a-1, 7, 'OUTR', $subhead );
$worksheet->merge_range( $a-2, 8, $a-1, 8, 'CONC', $subhead );
$worksheet->merge_range( $a-2, 9, $a-1, 9, 'TOTAL', $subhead );

$worksheet->merge_range( $a-3, 10, $a-3, 12, '%CONTR(ACTUAL)', $subhead );
$worksheet->merge_range( $a-2, 10, $a-2, 11, 'OUTR-CONC MIX', $subhead );
$worksheet->merge_range( $a-2, 12, $a-1, 12, 'TO TOTAL', $subhead );

$worksheet->write( $a-1, 10, 'OUTR', $subhead );
$worksheet->write( $a-1, 11, 'CONC', $subhead );

&heading_div;
&query_merch;

}


sub new_sheet_div{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(90);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
$worksheet->set_print_scale( 100 );
#$worksheet->fit_to_pages( 1, 1 );
$worksheet->set_margins( 0.001 );
$worksheet->conditional_formatting( 'F9:AK6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );

$worksheet->set_row( 3, undef, undef, 1, undef, undef ); #we hide this row
$worksheet->set_row( 4, undef, undef, 1, undef, undef ); #we hide this row

$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 23 );

$worksheet->set_column( 7, 7, 8 );

}

sub new_sheet_dly{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(90);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
$worksheet->set_print_scale( 100 );
#$worksheet->fit_to_pages( 1, 1 );
$worksheet->set_margins( 0.001 );
$worksheet->conditional_formatting( 'F9:AK6000',  { type => 'cell', criteria => '<', value => 0, format => $down } );


$worksheet->set_row( 3, undef, undef, 1, undef, undef ); #we hide this row
$worksheet->set_row( 4, undef, undef, 1, undef, undef ); #we hide this row
$worksheet->set_row( 5, undef, undef, 1, undef, undef ); #we hide this row

$worksheet->set_column( 0, 0, 3 );
$worksheet->set_column( 1, 4, undef, undef, 1 );

$worksheet->set_column( 5, 5, 7 );
$worksheet->set_column( 6, 6, 9 );
$worksheet->set_column( 7, 9, 14 );

}


sub heading_div {

$worksheet->merge_range( $a-2, 2, $a-1, 2, 'Type', $subhead );
$worksheet->merge_range( $a-2, 3, $a-1, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a-1, 4, 'Type', $subhead );
$worksheet->merge_range( $a-2, 5, $a-1, 5, 'Code', $subhead );
$worksheet->merge_range( $a-2, 6, $a-1, 6, 'Desc', $subhead );

}

sub heading_dly {

$worksheet->merge_range( $a-2, 5, $a-1, 5, 'Day #', $subhead );
$worksheet->merge_range( $a-2, 6, $a-1, 6, 'Date', $subhead );
$worksheet->merge_range( $a-2, 7, $a-2, 9, 'DAILY SALES', $subhead );
$worksheet->write( $a-1, 7, 'GENMERCH', $subhead );
$worksheet->write( $a-1, 8, 'SUPERMARKET', $subhead );
$worksheet->write( $a-1, 9, 'TOTAL', $subhead );

}


sub query_dept {

$temp_desc = 0;
$col_start = 7;
$day_counter = 1;
$day_count = $ns_st_date_key;

while ( $day_count <= $wk_en_date_key ){

	$total_net = 0;
			
	$sls1 = $dbh->prepare (qq{SELECT TO_CHAR(DATE_FLD, 'DD-MON') DATE_FLD, MERCH_GROUP_CODE_REV, MERCH_GROUP_DESC_REV, SUM(TARGET_SALE_VAL) AS TARGET, SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE
								  FROM METRO_IT_DLY_NEW_STORE
								  WHERE DATE_KEY = $day_count
								  GROUP BY DATE_FLD, MERCH_GROUP_CODE_REV, MERCH_GROUP_DESC_REV
								  ORDER BY MERCH_GROUP_CODE_REV
								 });								 
	$sls1->execute();

	while(my $s = $sls1->fetchrow_hashref()){
		$date_fld = $s->{DATE_FLD};
		$merch_group_code = $s->{MERCH_GROUP_CODE_REV};
		$merch_group_desc = $s->{MERCH_GROUP_DESC_REV};
		
		$sls2 = $dbh->prepare (qq{SELECT GROUP_CODE, GROUP_DESC, SUM(TARGET_SALE_VAL) AS TARGET, SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE
									 FROM METRO_IT_DLY_NEW_STORE 
									 WHERE MERCH_GROUP_CODE_REV = '$merch_group_code' AND DATE_KEY = $day_count
									 GROUP BY GROUP_CODE, GROUP_DESC
									 ORDER BY GROUP_CODE
								 });	
		$sls2->execute();
		
		$mgc_counter = $a;
		while(my $s = $sls2->fetchrow_hashref()){
			$group_code = $s->{GROUP_CODE};
			$group_desc = $s->{GROUP_DESC};
					
			$sls3 = $dbh->prepare (qq{SELECT DIVISION, DIVISION_DESC, SUM(TARGET_SALE_VAL) AS TARGET, SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE
										 FROM METRO_IT_DLY_NEW_STORE 
										 WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE_REV = '$merch_group_code' AND DATE_KEY = $day_count
										 GROUP BY DIVISION, DIVISION_DESC
										 ORDER BY DIVISION
										});
			$sls3->execute();
			
			$grp_counter = $a;
			while(my $s = $sls3->fetchrow_hashref()){
				$division = $s->{DIVISION};
				$division_desc = $s->{DIVISION_DESC};
				
				$sls4 = $dbh->prepare (qq{SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, SUM(TARGET_SALE_VAL) AS TARGET, SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE
											 FROM METRO_IT_DLY_NEW_STORE 
											 WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE_REV = '$merch_group_code' AND DIVISION = '$division' AND DATE_KEY = $day_count
											 GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC 
											 ORDER BY DEPARTMENT_CODE
											 });
				$sls4->execute();
				
				while(my $s = $sls4->fetchrow_hashref()){
					
					if ($temp_desc eq 0){
						$worksheet->write($a,5, $s->{DEPARTMENT_CODE},$desc);
						$worksheet->write($a,6, $s->{DEPARTMENT_DESC},$desc); }
					
					$worksheet->write($a,$col_start, $s->{NET_SALE},$border1);
						
					$counter++; 			
					$a++;
			
				}
				
				$worksheet->write($a,$col_start, $s->{NET_SALE},$bodyNum);
				
				if ($temp_desc eq 0){
					$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
					$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 ); }
				
				$counter = 0;
				$a++;
			}

			if($group_code ne 'JW'){
				$grp_net += $s->{NET_SALE};
			}
			
			$worksheet->write($a,$col_start, $s->{NET_SALE},$bodyNum);
			
			if ($temp_desc eq 0){
				$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
				$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 ); }

			$a++; #INCREMENT VARIABLE a
		}
		
		$total_net += $s->{NET_SALE};
		
		if ($merch_group_code eq 'DS' and $temp_desc eq 0){
			
			$worksheet->write($a,$col_start, $grp_net,$headNum);
			
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
			$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
		}
		
		elsif($merch_group_code eq 'SU' and $temp_desc eq 0){
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
			$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
		}
		
		elsif($merch_group_code eq 'Z_OT' and $temp_desc eq 0){
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Others', $headN );
			$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'OTHERS', $border2 );
		}
		
		$worksheet->write($a,$col_start, $s->{NET_SALE},$headNumber);
		
		$a++;
	}

	$worksheet->write($a,$col_start, $total_net,$headNumber);
		
	if ($temp_desc eq 0){
		$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN ); }
	
	$worksheet->write(8,$col_start, 'Day '.$day_counter,$subhead );
	$worksheet->write(9,$col_start, $date_fld,$subhead );

	$a = 10;
	$temp_desc = 1;
	$counter = 0;
	$col_start++;
	$day_count++;
	$day_counter++;

}

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

}

sub query_dly {

$counter = 0;
$day_counter = 1;
$day_count = $ns_st_date_key;

while ( $day_count <= $wk_en_date_key ){
			
	$sls = $dbh->prepare (qq{SELECT DATE_FLD, SUM(NET_SALE_DS) NET_SALE_DS, SUM(NET_SALE_SU) NET_SALE_SU FROM
								(SELECT DATE_FLD, SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE_DS, 0 AS NET_SALE_SU
								FROM METRO_IT_DLY_NEW_STORE
								WHERE MERCH_GROUP_CODE_REV = 'DS' AND DATE_KEY = $day_count 
								GROUP BY DATE_FLD
								UNION ALL
								SELECT DATE_FLD, 0 AS NET_SALE_DS, SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE_SU
								FROM METRO_IT_DLY_NEW_STORE
								WHERE MERCH_GROUP_CODE_REV = 'SU' AND DATE_KEY = $day_count
								GROUP BY DATE_FLD)
							GROUP BY DATE_FLD
							});
	$sls->execute();
				
	while(my $s = $sls->fetchrow_hashref()){
		$worksheet->write($a,5, 'Day '.$day_counter,$desc);
		$worksheet->write($a,6, $s->{DATE_FLD},$desc);
		$worksheet->write($a,7, $s->{NET_SALE_DS},$border1); 			
		$worksheet->write($a,8, $s->{NET_SALE_SU},$border1);
		$worksheet->write($a,9, $s->{NET_SALE_DS}+$s->{NET_SALE_SU},$border1);
								
		$a++;
		$counter++;
			
	}

	$day_count++;
	$day_counter++;

}

$day_counter -= 1;

$worksheet->merge_range($a,5,$a,6, 'TOTAL',$headN);
$worksheet->write($a,7, '=SUM('. xl_rowcol_to_cell( $a-$counter, 7 ). ':' . xl_rowcol_to_cell( $a-1, 7 ) . ')',$headNumber); 			
$worksheet->write($a,8, '=SUM('. xl_rowcol_to_cell( $a-$counter, 8 ). ':' . xl_rowcol_to_cell( $a-1, 8 ) . ')',$headNumber);
$worksheet->write($a,9, '=SUM('. xl_rowcol_to_cell( $a-$counter, 9 ). ':' . xl_rowcol_to_cell( $a-1, 9 ) . ')',$headNumber);

$worksheet->merge_range($a+1,5,$a+1,6, 'Ave Daily Sales',$headN);
$worksheet->write($a+1,7, '=(SUM('. xl_rowcol_to_cell( $a-$counter, 7 ). ':' . xl_rowcol_to_cell( $a-1, 7 ) . '))/'.$day_counter,$headNumber); 			
$worksheet->write($a+1,8, '=(SUM('. xl_rowcol_to_cell( $a-$counter, 8 ). ':' . xl_rowcol_to_cell( $a-1, 8 ) . '))/'.$day_counter,$headNumber);
$worksheet->write($a+1,9, '=(SUM('. xl_rowcol_to_cell( $a-$counter, 9 ). ':' . xl_rowcol_to_cell( $a-1, 9 ) . '))/'.$day_counter,$headNumber);

$worksheet->merge_range($a+2,5,$a+2,6, 'Ave Daily Sales Plan',$headN);
$worksheet->write($a+2,7, '',$headNumber); 			
$worksheet->write($a+2,8, '',$headNumber);
$worksheet->write($a+2,9, '',$headNumber);

$worksheet->merge_range($a+3,5,$a+3,6, 'Var from Plan',$headN);
$worksheet->write($a+3,7, '',$headNumber); 			
$worksheet->write($a+3,8, '',$headNumber);
$worksheet->write($a+3,9, '',$headNumber);

$worksheet->merge_range($a+4,5,$a+4,6, 'Actual Mix',$headN);
$worksheet->write($a+4,7, '=(SUM('. xl_rowcol_to_cell( $a-$counter, 7 ). ':' . xl_rowcol_to_cell( $a-1, 7 ) . '))/(SUM('. xl_rowcol_to_cell( $a-$counter, 9 ). ':' . xl_rowcol_to_cell( $a-1, 9 ) . '))',$headPct); 			
$worksheet->write($a+4,8, '=(SUM('. xl_rowcol_to_cell( $a-$counter, 8 ). ':' . xl_rowcol_to_cell( $a-1, 8 ) . '))/(SUM('. xl_rowcol_to_cell( $a-$counter, 9 ). ':' . xl_rowcol_to_cell( $a-1, 9 ) . '))',$headPct);
$worksheet->write($a+4,9, '100 %',$headPct);

$sls->finish();

}

sub query_merch {

$col_start = 7;
$mg_total = 0;
	
$sls1 = $dbh->prepare (qq{SELECT MERCH_GROUP_CODE_REV, MERCH_GROUP_DESC_REV, 
								  SUM(NET_SALE_OTR) AS NET_SALE_OTR, 
								  SUM(NET_SALE_CON) AS NET_SALE_CON, 
								  SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE,
CASE WHEN (SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0))) = 0 THEN NULL ELSE ROUND(((SUM(NVL(NET_SALE_OTR,0)))/(SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)))),1) END AS OTR_CONT,
CASE WHEN (SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0))) = 0 THEN NULL ELSE ROUND(((SUM(NVL(NET_SALE_CON,0)))/(SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)))),1) END AS CON_CONT
							FROM METRO_IT_DLY_NEW_STORE
							GROUP BY MERCH_GROUP_CODE_REV, MERCH_GROUP_DESC_REV
							ORDER BY MERCH_GROUP_CODE_REV
							});								 
$sls1->execute();

	while(my $s = $sls1->fetchrow_hashref()){
		$date_fld = $s->{DATE_FLD};
		$merch_group_code = $s->{MERCH_GROUP_CODE_REV};
		$merch_group_desc = $s->{MERCH_GROUP_DESC_REV};
		$mg_total = $s->{NET_SALE};
		
		$sls2 = $dbh->prepare (qq{SELECT GROUP_CODE, GROUP_DESC, 
									SUM(NET_SALE_OTR) AS NET_SALE_OTR, 
									SUM(NET_SALE_CON) AS NET_SALE_CON, 
									SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE,
CASE WHEN (SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0))) = 0 THEN NULL ELSE ROUND(((SUM(NVL(NET_SALE_OTR,0)))/(SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)))),1) END AS OTR_CONT,
CASE WHEN (SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0))) = 0 THEN NULL ELSE ROUND(((SUM(NVL(NET_SALE_CON,0)))/(SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)))),1) END AS CON_CONT
								FROM METRO_IT_DLY_NEW_STORE 
								WHERE MERCH_GROUP_CODE_REV = '$merch_group_code'
								GROUP BY GROUP_CODE, GROUP_DESC
								ORDER BY GROUP_CODE
								 });	
		$sls2->execute();
		
		$mgc_counter = $a;
		while(my $s = $sls2->fetchrow_hashref()){
			$group_code = $s->{GROUP_CODE};
			$group_desc = $s->{GROUP_DESC};
					
			$sls3 = $dbh->prepare (qq{SELECT DIVISION, DIVISION_DESC, 
										SUM(NET_SALE_OTR) AS NET_SALE_OTR, 
										SUM(NET_SALE_CON) AS NET_SALE_CON, 
										SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE,
CASE WHEN (SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0))) = 0 THEN NULL ELSE ROUND(((SUM(NVL(NET_SALE_OTR,0)))/(SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)))),1) END AS OTR_CONT,
CASE WHEN (SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0))) = 0 THEN NULL ELSE ROUND(((SUM(NVL(NET_SALE_CON,0)))/(SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)))),1) END AS CON_CONT
									FROM METRO_IT_DLY_NEW_STORE 
									WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE_REV = '$merch_group_code'
									GROUP BY DIVISION, DIVISION_DESC
									ORDER BY DIVISION
									});
			$sls3->execute();
			
			$grp_counter = $a;
			while(my $s = $sls3->fetchrow_hashref()){
				$division = $s->{DIVISION};
				$division_desc = $s->{DIVISION_DESC};
				
				$sls4 = $dbh->prepare (qq{SELECT DEPARTMENT_CODE, DEPARTMENT_DESC, 
											SUM(NET_SALE_OTR) AS NET_SALE_OTR, 
											SUM(NET_SALE_CON) AS NET_SALE_CON, 
											SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)) AS NET_SALE,
CASE WHEN (SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0))) = 0 THEN NULL ELSE ROUND(((SUM(NVL(NET_SALE_OTR,0)))/(SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)))),1) END AS OTR_CONT,
CASE WHEN (SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0))) = 0 THEN NULL ELSE ROUND(((SUM(NVL(NET_SALE_CON,0)))/(SUM(NVL(NET_SALE_OTR,0)+NVL(NET_SALE_CON,0)))),1) END AS CON_CONT
										FROM METRO_IT_DLY_NEW_STORE 
										WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE_REV = '$merch_group_code' AND DIVISION = '$division'
										GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC 
										ORDER BY DEPARTMENT_CODE
										});
				$sls4->execute();
				
				while(my $s = $sls4->fetchrow_hashref()){
					$worksheet->write($a,5, $s->{DEPARTMENT_CODE},$desc);
					$worksheet->write($a,6, $s->{DEPARTMENT_DESC},$desc); 
					
					$worksheet->write($a,$col_start, $s->{NET_SALE_OTR},$border1);
					$worksheet->write($a,$col_start+1, $s->{NET_SALE_CON},$border1);
					$worksheet->write($a,$col_start+2, $s->{NET_SALE},$border1);
					$worksheet->write($a,$col_start+3, $s->{OTR_CONT},$subt);
					$worksheet->write($a,$col_start+4, $s->{CON_CONT},$subt);
					if ($mg_total eq 0){
						$worksheet->write($a,$col_start+5, '',$subt);}
					else{
						$worksheet->write($a,$col_start+5, $s->{NET_SALE}/$mg_total,$subt);}
						
					$counter++; 			
					$a++;
			
				}
				
				$worksheet->write($a,$col_start, $s->{NET_SALE_OTR},$bodyNum);
				$worksheet->write($a,$col_start+1, $s->{NET_SALE_CON},$bodyNum);
				$worksheet->write($a,$col_start+2, $s->{NET_SALE},$bodyNum);
				$worksheet->write($a,$col_start+3, $s->{OTR_CONT},$bodyPct);
				$worksheet->write($a,$col_start+4, $s->{CON_CONT},$bodyPct);
				if ($mg_total eq 0){
					$worksheet->write($a,$col_start+5, '',$bodyPct);}
				else{
					$worksheet->write($a,$col_start+5, $s->{NET_SALE}/$mg_total,$bodyPct);}
				
				$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
				$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
				
				$counter = 0;
				$a++;
			}

			if($group_code ne 'JW'){
				$grp_otr += $s->{NET_SALE_OTR};
				$grp_con += $s->{NET_SALE_CON};
				$grp_net += $s->{NET_SALE};
			}
			
			$worksheet->write($a,$col_start, $s->{NET_SALE_OTR},$bodyNum);
			$worksheet->write($a,$col_start+1, $s->{NET_SALE_CON},$bodyNum);
			$worksheet->write($a,$col_start+2, $s->{NET_SALE},$bodyNum);
			$worksheet->write($a,$col_start+3, $s->{OTR_CONT},$bodyPct);
			$worksheet->write($a,$col_start+4, $s->{CON_CONT},$bodyPct);
			if($mg_total eq 0){
				$worksheet->write($a,$col_start+5, '',$bodyPct);}
			else{
				$worksheet->write($a,$col_start+5, $s->{NET_SALE}/$mg_total,$bodyPct);}
			
			
			$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );

			$a++; #INCREMENT VARIABLE a
		}
		
		$total_otr += $s->{NET_SALE_OTR};
		$total_con += $s->{NET_SALE_CON};
		$total_net += $s->{NET_SALE};
		
		if ($merch_group_code eq 'DS'){
			
			$worksheet->write($a,$col_start, $grp_otr,$headNum);
			$worksheet->write($a,$col_start+1, $grp_con,$headNum);
			$worksheet->write($a,$col_start+2, $grp_net,$headNum);
			
			if ($grp_net eq 0) {
				$worksheet->write($a,$col_start+3, '',$headDPct);
				$worksheet->write($a,$col_start+4, '',$headDPct);
				$worksheet->write($a,$col_start+5, '',$headDPct);}
			else {
				$worksheet->write($a,$col_start+3, $grp_otr/$grp_net,$headDPct);
				$worksheet->write($a,$col_start+4, $grp_con/$grp_net,$headDPct);
				if($mg_total eq 0){
					$worksheet->write($a,$col_start+5, '',$headDPct);}
				else{
					$worksheet->write($a,$col_start+5, $grp_net/$mg_total,$headDPct);}}		
			
			
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
		
		$worksheet->write($a,$col_start, $s->{NET_SALE_OTR},$headNumber);
		$worksheet->write($a,$col_start+1, $s->{NET_SALE_CON},$headNumber);
		$worksheet->write($a,$col_start+2, $s->{NET_SALE},$headNumber);
		$worksheet->write($a,$col_start+3, $s->{OTR_CONT},$headPct);
		$worksheet->write($a,$col_start+4, $s->{CON_CONT},$headPct);
		if($mg_total eq 0) {
			$worksheet->write($a,$col_start+5, '',$headPct);}
		else{
			$worksheet->write($a,$col_start+5, $s->{NET_SALE}/$mg_total,$headPct);}
		
		
		$a++;
	}

	$worksheet->write($a,$col_start, $total_net,$headNumber);
	
	$worksheet->write($a,$col_start, $total_otr,$headNumber);
	$worksheet->write($a,$col_start+1, $total_con,$headNumber);
	$worksheet->write($a,$col_start+2, $total_net,$headNumber);
	
	if ($total_net eq 0){
		$worksheet->write($a,$col_start+3, '',$headPct);
		$worksheet->write($a,$col_start+4, '',$headPct);
		$worksheet->write($a,$col_start+5, '',$headPct);}
	else{
		$worksheet->write($a,$col_start+3, $total_otr/$total_net,$headPct);
		$worksheet->write($a,$col_start+4, $total_con/$total_net,$headPct);
		$worksheet->write($a,$col_start+5, $total_net/$total_net,$headPct);}
		
	$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

}


sub generate_csv {

my $truncate = $dbh->prepare( qq{ 
TRUNCATE TABLE METRO_IT_DLY_NEW_STORE 
});
$truncate->execute();

print "Done truncating METRO_IT_DLY_NEW_STORE... \nPreparing to Insert... \n";

$test = qq{ 
INSERT INTO METRO_IT_DLY_NEW_STORE
(DATE_KEY, DATE_FLD, STORE_FORMAT, STORE_FORMAT_DESC, STORE_CODE, STORE, STORE_DESCRIPTION, 
MERCH_GROUP_CODE_REV, MERCH_GROUP_DESC_REV, MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC, DEPARTMENT_CODE, DEPARTMENT_DESC, 
NEW_FLG, 
MATURED_FLG, 
TARGET_SALE_VAL, 
NET_SALE_OTR, 
NET_SALE_CON)

SELECT DATE_KEY, DATE_FLD, STORE_FORMAT, STORE_FORMAT_DESC, STORE_CODE, STORE, STORE_DESCRIPTION, MERCH_GROUP_CODE_REV, MERCH_GROUP_DESC_REV, MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC, DEPARTMENT_CODE, DEPARTMENT_DESC, 
case when store_code in ('3009','3012','6004','6009','6012') then 1 else 0 end as NEW_FLG, 
case when store_code in ('3009','3012','6004','6009','6012') then 0 else 1 end as MATURED_FLG, 
TARGET_SALE_VAL, NET_SALE_OTR, NET_SALE_CON
FROM (
SELECT 
BASE.DATE_KEY, 
BASE.DATE_FLD, 
BASE.STORE_FORMAT, 
BASE.STORE_FORMAT_DESC, 
BASE.STORE_CODE, 
CASE WHEN BASE.STORE_CODE IN ('2012', '2013', '3009', '4004', '3010', '3011', '2001W', '3012') THEN 'SU' || BASE.STORE_CODE 
     WHEN BASE.STORE_CODE = '2223' THEN 'DS' || BASE.STORE_CODE 
	 ELSE BASE.MERCH_GROUP_CODE || BASE.STORE_CODE END AS STORE,	 
UPPER(BASE.STORE_DESCRIPTION) AS STORE_DESCRIPTION,
CASE WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 9000) THEN 'Z_OT'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8500) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE = 8040) THEN 'SU'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE != 8040) THEN 'DS'
ELSE BASE.MERCH_GROUP_CODE END AS MERCH_GROUP_CODE_REV,
CASE WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 9000) THEN 'OTHERS'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8500) THEN 'SUPERMARKET'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE = 8040) THEN 'SUPERMARKET'
     WHEN (BASE.MERCH_GROUP_CODE = 'OT' AND BASE.DIVISION = 8000 AND BASE.DEPARTMENT_CODE != 8040) THEN 'DEPARTMENT STORE'
ELSE BASE.MERCH_GROUP_DESC END AS MERCH_GROUP_DESC_REV,
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
SUM(WTD.TARGET_SALE_VAL) TARGET_SALE_VAL, SUM(WTD.NET_SALE_OTR) NET_SALE_OTR, SUM(WTD.NET_SALE_CON) NET_SALE_CON
, SUM(NVL(WTD.TARGET_SALE_VAL,0)+NVL(WTD.NET_SALE_OTR,0)+NVL(WTD.NET_SALE_CON,0)) SALE_CHECK
FROM

(SELECT D.DATE_KEY, D.DATE_FLD, S.STORE_KEY, S.STORE_FORMAT, S.STORE_FORMAT_DESC, 
	CASE WHEN S.STORE_CODE IN ('4002') THEN '2001W'
		ELSE S.STORE_CODE END AS STORE_CODE,
		UPPER(S.STORE_DESCRIPTION) AS STORE_DESCRIPTION, 
	M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.GROUP_CODE, M.GROUP_DESC, M.DIVISION, M.DIVISION_DESC, M.DEPARTMENT_CODE, M.DEPARTMENT_DESC,S.NEW_FLG, S.MATURED_FLG
FROM
	(SELECT DATE_KEY, DATE_FLD
			FROM DIM_DATE 
			WHERE DATE_KEY BETWEEN $ns_st_date_key AND $wk_en_date_key)D,
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
GROUP BY D.DATE_KEY, D.DATE_FLD, S.STORE_KEY, S.STORE_FORMAT, S.STORE_FORMAT_DESC, 
	CASE WHEN S.STORE_CODE IN ('4002') THEN '2001W'
		ELSE S.STORE_CODE END,
		UPPER(S.STORE_DESCRIPTION), 
	M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.GROUP_CODE, M.GROUP_DESC, M.DIVISION, M.DIVISION_DESC, M.DEPARTMENT_CODE, M.DEPARTMENT_DESC,S.NEW_FLG, S.MATURED_FLG)BASE
	
LEFT JOIN

(SELECT 
AGG_MLY_STR_DEPT_TARGET.DATE_KEY, 
DIM_STORE.STORE_KEY, 
DIM_STORE.STORE_FORMAT, 
CASE WHEN DIM_STORE.STORE_CODE IN ('4002') THEN '2001W'
	ELSE DIM_STORE.STORE_CODE END AS STORE_CODE,
DIM_SUB_DEPT.MERCH_GROUP_CODE,
DIM_SUB_DEPT.GROUP_CODE,
DIM_SUB_DEPT.DIVISION,
DIM_SUB_DEPT.DEPARTMENT_CODE,
NVL((SUM(TARGET_SALE_VAL))/1000,0) TARGET_SALE_VAL,  
NVL((SUM((NVL(SALE_NET_VAL_OTR,0))-(NVL(SALE_TOT_TAX_VAL_OTR,0))-(NVL(SALE_TOT_DISC_VAL_OTR,0))))/1000,0) NET_SALE_OTR,
NVL((SUM((NVL(SALE_NET_VAL_CON,0))-(NVL(SALE_TOT_TAX_VAL_CON,0))-(NVL(SALE_TOT_DISC_VAL_CON,0))))/1000,0) NET_SALE_CON
FROM (	
	SELECT TBL.DATE_KEY, TBL.STORE_KEY STORE_KEY, TBL.DS_KEY DS_KEY, TBL.STORE_CODE STORE_CODE, TY.SALE_NET_VAL_OTR, TY.SALE_TOT_TAX_VAL_OTR, TY.SALE_TOT_DISC_VAL_OTR, LY.SALE_NET_VAL_CON, LY.SALE_TOT_TAX_VAL_CON, LY.SALE_TOT_DISC_VAL_CON,
0 AS TARGET_SALE_VAL, 0 AS TARGET_SALE_VAL_LY, 0 AS TARGET_SALE_VAT, 0 AS TARGET_SALE_VAT_LY 	FROM
		(SELECT D.DATE_KEY, S.STORE_KEY, S.STORE_CODE, M.DS_KEY
		FROM
			(SELECT DATE_KEY, DATE_FLD
			FROM DIM_DATE 
			WHERE DATE_KEY BETWEEN $ns_st_date_key AND $wk_en_date_key )D,
			(SELECT STORE_KEY, STORE_CODE
			FROM DIM_STORE 
			WHERE ACTIVE = 1 AND STORE_FORMAT IN (1, 2, 3, 4, 5))S,
			(SELECT D.DS_KEY, D.MERCH_GROUP_CODE, D.MERCH_GROUP_DESC, D.GROUP_CODE, D.GROUP_DESC, D.DIVISION, D.DIVISION_DESC, D.DEPARTMENT_CODE, D.DEPARTMENT_DESC
			FROM DIM_MERCHANDISE M 
				JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION 
					AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE )M
		GROUP BY D.DATE_KEY, S.STORE_KEY, S.STORE_CODE, M.DS_KEY)TBL
		LEFT JOIN
		
		(SELECT DATE_KEY, 
		STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE,  
		SUM(SALE_NET_VAL) AS SALE_NET_VAL_OTR, 
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_OTR, 
		SUM(SALE_TOT_DISC_VAL) SALE_TOT_DISC_VAL_OTR
		FROM AGG_DLY_STR_PROD AGG 
			INNER JOIN DIM_PRODUCT M ON AGG.PRODUCT_KEY = M.PRODUCT_KEY AND M.PRODUCT_CATEGORY = 0
			INNER JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
		WHERE DATE_KEY BETWEEN $ns_st_date_key AND $wk_en_date_key 
		GROUP BY DATE_KEY, STORE_KEY, DS_KEY, STORE_CODE)TY
		
		ON TBL.DATE_KEY = TY.DATE_KEY AND TBL.STORE_KEY = TY.STORE_KEY AND TBL.STORE_CODE = TY.STORE_CODE AND TBL.DS_KEY = TY.DS_KEY
		LEFT JOIN
		
		(SELECT DATE_KEY, 
		STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE,  
		SUM(SALE_NET_VAL) AS SALE_NET_VAL_CON, 
		SUM(SALE_TOT_TAX_VAL) SALE_TOT_TAX_VAL_CON, 
		SUM(SALE_TOT_DISC_VAL) SALE_TOT_DISC_VAL_CON
		FROM AGG_DLY_STR_PROD AGG 
			INNER JOIN DIM_PRODUCT M ON AGG.PRODUCT_KEY = M.PRODUCT_KEY AND M.PRODUCT_CATEGORY = 2 
			INNER JOIN DIM_SUB_DEPT D ON M.MERCH_GROUP_CODE = D.MERCH_GROUP_CODE AND M.GROUP_CODE = D.GROUP_CODE AND M.GROUP2_CODE = D.GROUP2_CODE AND M.DIVISION = D.DIVISION AND M.DEPARTMENT_CODE = D.DEPARTMENT_CODE AND M.SUB_DEPARTMENT_CODE = D.SUB_DEPARTMENT_CODE
		WHERE DATE_KEY BETWEEN $ns_st_date_key AND $wk_en_date_key
		GROUP BY DATE_KEY, STORE_KEY, DS_KEY, STORE_CODE)LY 
		
		ON TBL.DATE_KEY = LY.DATE_KEY AND TBL.STORE_KEY = LY.STORE_KEY AND TBL.STORE_CODE = LY.STORE_CODE AND TBL.DS_KEY = LY.DS_KEY 
	
	UNION ALL 
	
	SELECT A.DATE_KEY, STORE_KEY, DS_KEY, CAST(STORE_CODE AS VARCHAR2(20)) AS STORE_CODE, 
		0 AS SALE_NET_VAL_OTR, 0 AS SALE_TOT_TAX_VAL_OTR, 0 AS SALE_TOT_DISC_VAL_OTR, 0 AS SALE_NET_VAL_CON, 0 AS SALE_TOT_TAX_VAL_CON, 0 AS SALE_TOT_DISC_VAL_CON,		
		SUM (TARGET_SALE_VAL) AS TARGET_SALE_VAL, 
		SUM (TARGET_SALE_VAL_LY) AS TARGET_SALE_VAL_LY, 
		SUM (TARGET_SALE_VAT) AS TARGET_SALE_VAT, 
		SUM (TARGET_SALE_VAT_LY) AS TARGET_SALE_VAT_LY 
	FROM FCT_TARGET A JOIN DIM_DATE_PRL DP ON A.DATE_KEY = DP.DATE_KEY 
	WHERE A.DATE_KEY BETWEEN $ns_st_date_key AND $wk_en_date_key
	GROUP BY A.DATE_KEY, STORE_KEY, STORE_CODE, DS_KEY 		
		
		) AGG_MLY_STR_DEPT_TARGET,DIM_STORE,DIM_SUB_DEPT,DIM_DATE
WHERE DIM_STORE.ACTIVE = 1 AND DIM_STORE.STORE_FORMAT IN (1, 2, 3, 4, 5) AND AGG_MLY_STR_DEPT_TARGET.STORE_KEY=DIM_STORE.STORE_KEY AND AGG_MLY_STR_DEPT_TARGET.DS_KEY=DIM_SUB_DEPT.DS_KEY AND AGG_MLY_STR_DEPT_TARGET.DATE_KEY = DIM_DATE.DATE_KEY
GROUP BY 
	AGG_MLY_STR_DEPT_TARGET.DATE_KEY, 
	DIM_STORE.STORE_KEY, DIM_STORE.STORE_FORMAT, 
	CASE WHEN DIM_STORE.STORE_CODE IN ('4002') THEN '2001W'
		ELSE DIM_STORE.STORE_CODE END, 
	DIM_SUB_DEPT.MERCH_GROUP_CODE, DIM_SUB_DEPT.GROUP_CODE, DIM_SUB_DEPT.DIVISION, DIM_SUB_DEPT.DEPARTMENT_CODE
)WTD

ON BASE.DATE_KEY = WTD.DATE_KEY AND BASE.STORE_KEY = WTD.STORE_KEY AND BASE.STORE_FORMAT = WTD.STORE_FORMAT AND BASE.STORE_CODE = WTD.STORE_CODE AND BASE.MERCH_GROUP_CODE = WTD.MERCH_GROUP_CODE AND BASE.GROUP_CODE = WTD.GROUP_CODE AND BASE.DIVISION = WTD.DIVISION AND BASE.DEPARTMENT_CODE = WTD.DEPARTMENT_CODE

GROUP BY
BASE.DATE_KEY, 
BASE.DATE_FLD,
BASE.STORE_FORMAT, 
BASE.STORE_FORMAT_DESC, 
BASE.STORE_CODE, 
CASE WHEN BASE.STORE_CODE IN ('2012', '2013', '3009', '4004', '3010', '3011', '2001W', '3012') THEN 'SU' || BASE.STORE_CODE
     WHEN BASE.STORE_CODE = '2223' THEN 'DS' || BASE.STORE_CODE 
	 ELSE BASE.MERCH_GROUP_CODE || BASE.STORE_CODE END,	 
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
)
WHERE STORE_CODE = '3010'
};

my $sth = $dbh->prepare ($test);
$sth->execute;
 
$sth->finish();
$dbh->commit;

print "Done with Insert...";

}


sub mail_grp1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1 ) = @ARGV;

$to = ' emily.silverio@metrogaisano.com, limuel.ulanday@metrogaisano.com, maricel.tamala@metrogaisano.com, fili.mercado@metrogaisano.com, chit.lazaro@metrogaisano.com, karan.malani@metrogaisano.com, lia.chipeco@metrogaisano.com, luz.bitang@metrogaisano.com ';

$cc = ' arthur.emmanuel@metrogaisano.com, rex.cabanilla@metrogaisano.com, annalyn.conde@metrogaisano.com ';

$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, cham.burgos@metrogaisano.com';

#$to = ' kent.mamalias@metrogaisano.com ';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';

$subject = 'New Store Daily Sales Performance as of ' . $as_of;

$msgbody_file = 'message_BI_new_store.txt';

$attachment_file_1 = "New Store Daily Sales Performance - Summary (as of $as_of) v1.xlsx";

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

sub mail_grp1_lateposted {

my $table = 'late_posted_'.$day . '_' . $month_to_text . '_' .$year . '.csv';

my $dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
	or die $DBI::errstr;
 
my $sth = $dbh_csv->prepare(qq{SELECT STORE, STORE_NAME, MERCH_GROUP_CODE_REV, SUM(VALUE) VALUE 
								FROM $table  WHERE STORE = '3010'
								GROUP BY STORE, STORE_NAME, MERCH_GROUP_CODE_REV 
								ORDER BY STORE, MERCH_GROUP_CODE_REV});

$sth->execute() or die "Failed to execute query - " . $dbh_csv->errstr;

my $table = HTML::Table::FromDatabase->new( -sth => $sth );

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1 ) = @ARGV;

$to = ' emily.silverio@metrogaisano.com, limuel.ulanday@metrogaisano.com, maricel.tamala@metrogaisano.com, fili.mercado@metrogaisano.com, chit.lazaro@metrogaisano.com, karan.malani@metrogaisano.com, lia.chipeco@metrogaisano.com, luz.bitang@metrogaisano.com ';

$cc = ' arthur.emmanuel@metrogaisano.com, rex.cabanilla@metrogaisano.com, annalyn.conde@metrogaisano.com ';

$bcc = ' kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, cham.burgos@metrogaisano.com';

#$to = ' kent.mamalias@metrogaisano.com ';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';

$subject = 'New Store Daily Sales Performance as of ' . $as_of;

$msgbody_file = 'message_BI_new_store.txt';

$attachment_file_1 = "New Store Daily Sales Performance - Summary (as of $as_of) v1.xlsx";

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
Content-Type: text/html; charset="iso-8859-1"
Content-Transfer-Encoding: quoted-printable

<html>
Hi All, <br> <br>
Please see attached Daily Sales Summary Report(in excel and pdf format). <br> <br>

Please be advised that the DSP as of $as_of is incomplete for the list of stores below due to delayed sales posting. The IT team is currently reviewing the  process and finding ways to further stabilize the DSP. An updated report will be sent tomorrow after found complete and accurate. <br> <br>

$table <br>

If you need assistance, kindly email arcbi.support&#64;metrogaisano.com. <br> <br>

Regards, <br>
ARC BI Support <p>
</html>

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

sub mail_external {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1 ) = @ARGV;

$to = ' artemm12@aol.com, frankgaisano@gmail.com ';

#$to = ' kent.mamalias@gmail.com ';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';	

$subject = 'New Store Daily Sales Performance as of ' . $as_of;

$msgbody_file = 'message_BI_new_store.txt';

$attachment_file_1 = "New Store Daily Sales Performance - Summary (as of $as_of) v1.xlsx";

my $msgbody = read_file( $msgbody_file );

my $attachment_data_1 = encode_base64( read_file( $attachment_file_1, 1 ));

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

sub mail_external_lateposted {

my $table = 'late_posted_'.$day . '_' . $month_to_text . '_' .$year . '.csv';

my $dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
	or die $DBI::errstr;
 
my $sth = $dbh_csv->prepare(qq{SELECT STORE, STORE_NAME, MERCH_GROUP_CODE_REV, SUM(VALUE) VALUE 
								FROM $table  WHERE STORE = '3010'
								GROUP BY STORE, STORE_NAME, MERCH_GROUP_CODE_REV 
								ORDER BY STORE, MERCH_GROUP_CODE_REV});

$sth->execute() or die "Failed to execute query - " . $dbh_csv->errstr;

my $table = HTML::Table::FromDatabase->new( -sth => $sth );

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file_1 ) = @ARGV;

$to = ' artemm12@aol.com, frankgaisano@gmail.com ';

#$to = ' kent.mamalias@gmail.com ';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';	

$subject = 'Daily Sales Performance as of ' . $as_of;

$msgbody_file = 'message_BI_new_store.txt';

$attachment_file_1 = "New Store Daily Sales Performance - Summary (as of $as_of) v1.xlsx";

my $msgbody = read_file( $msgbody_file );

my $attachment_data_1 = encode_base64( read_file( $attachment_file_1, 1 ));

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
Please see attached Daily Sales Summary Report(in excel and pdf format). <br> <br>

Please be advised that the DSP as of $as_of is incomplete for the list of stores below due to delayed sales posting. The IT team is currently reviewing the  process and finding ways to further stabilize the DSP. An updated report will be sent tomorrow after found complete and accurate. <br> <br>

$table <br>

If you need assistance, kindly email arcbi.support&#64;metrogaisano.com. <br> <br>

Regards, <br>
ARC BI Support <p>
</html>

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






