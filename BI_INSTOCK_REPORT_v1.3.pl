use DBI;
use DBD::Oracle qw(:ora_types);
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use Text::CSV_XS;
use Win32::Job;
use Archive::Zip qw( :ERROR_CODES :CONSTANTS );
use Getopt::Long;
use IO::File;
use MIME::QuotedPrint;
use MIME::Base64;
use Mail::Sendmail;
use Date::Calc qw( Today Add_Delta_Days Month_to_Text);
use DBConnector;

($year,$month,$day) = Today();
$month_to_text = Month_to_Text($month);

my $workbook = Excel::Writer::XLSX->new("Replenishment In-stock v1.3.xlsx");
my $bold = $workbook->add_format( bold => 1, size => 14 );
my $bold1 = $workbook->add_format( bold => 1, size => 16 );
my $script = $workbook->add_format( size => 8, italic => 1 );
my $bold2 = $workbook->add_format( size => 11 );
my $border1 = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3 );
my $border2 = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', rotation => 90, text_wrap =>1, size => 10, shrink => 1 );
my $code = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10 );
my $desc = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10 );
my $ponkan = $workbook->set_custom_color( 53, 254, 238, 230);
my $abo = $workbook->set_custom_color( 16, 220, 218, 219);
my $sky = $workbook->set_custom_color( 12, 205, 225, 255);
my $pula = $workbook->set_custom_color( 10, 255, 189, 189);
my $lumot = $workbook->set_custom_color( 17, 196, 189, 151);
my $comp = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10,  bg_color => $lumot, bold => 1 );
my $all = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10,  bg_color => $abo, bold => 1 );
my $headN = $workbook->add_format( border => 1, align => 'center', valign => 'center', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
my $headD = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 9, bg_color => $abo, bold => 1 );
my $headPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
my $headNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3, bg_color => $abo, bold => 1 );
my $headNumber = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 3, bg_color => $abo, bold => 1 );
my $head = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
my $subhead = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, bg_color => $ponkan, bold => 1 );
my $bodyN = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
my $bodyPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
my $bodyNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 3,  bold => 1);
my $body = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
my $subt = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 9);
my $down = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => 9, bg_color => $pula );

printf "IN STOCK REPORT \n";

&new_sheet($sheet = "Summary-Store");
&call_str_merchandise;

&new_sheet($sheet = "Summary-Division");
&call_summary_division;

&new_sheet($sheet = "Store", $comment = "STORES");		
&call_vis;
&call_luz;

&new_sheet($sheet = "Warehouse", $comment = "WAREHOUSE");
&call_wh;

$workbook->close();

&mail1;	
&mail2;	
&mail3;	

exit;
 
#function call
sub call_vis {

$a = 0, $col = 7, $stopper = 0, $tot_vis = 0, $tot_luz = 0, $tot_wh = 0, $tot = 0;

&heading_orig;

$st = $dbh->prepare (qq{
			SELECT DISTINCT STORE_CODE
			FROM METRO_IT_INSTOCK_DEPT
			WHERE UPPER(REGION) LIKE '%VISAYAS%' AND (  (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0
			ORDER BY 1
			});								 
$st->execute();

while(my $s = $st->fetchrow_hashref()){

	$a += 6, $counter = 0, $loc_pt = $a-2;

	$worksheet->set_column( $col, $col+1, 7 );
	$worksheet->set_column( $col+2, $col+2, 5 );

	foreach my $i ("COUNT", "IN STOCK", "%") {
		$worksheet->write($a-1, $col++, $i, $subhead);
	}

	$col -= 3;

	&query_dept_store( $store = $s->{STORE_CODE} );

	$a = 0, $counter = 0, $stopper = 1, $col += 3;

}
	
$st->finish();

#========== TOTAL VISAYAS STORES

$a += 6, $counter = 0;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $headN);
}

$col -= 3;

&query_dept_store( $tot_vis = 1 );

$a = 0, $counter = 0, $col += 3, $tot_vis = 0;

}

sub call_luz {

$st = $dbh->prepare (qq{
			SELECT DISTINCT STORE_CODE
			FROM METRO_IT_INSTOCK_DEPT
			WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY')
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0
			ORDER BY 1
			});								 
$st->execute();

while(my $s = $st->fetchrow_hashref()){

	$a += 6, $counter = 0, $loc_pt = $a-2;

	$worksheet->set_column( $col, $col+1, 7 );
	$worksheet->set_column( $col+2, $col+2, 5 );

	foreach my $i ("COUNT", "IN STOCK", "%") {
		$worksheet->write($a-1, $col++, $i, $subhead);
	}

	$col -= 3;

	&query_dept_store( $store = $s->{STORE_CODE} );

	$a = 0, $counter = 0, $stopper = 1, $col += 3;

}
	
$st->finish();

#========== TOTAL LUZON STORES

$a += 6, $counter = 0;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $headN);
}

$col -= 3;

&query_dept_store( $tot_luz = 1 );

$a = 0, $counter = 0, $col += 3, $tot_luz = 0;

#========== TOTAL VISAYAS + LUZON STORES

$a += 6, $counter = 0;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $headN);
}

$col -= 3;

&query_dept_store( $tot = 1 );

$a = 0, $counter = 0, $col += 3, $tot = 0;

}

sub call_wh {

$a = 0, $col = 7, $stopper = 0;

&heading_orig;

$st = $dbh->prepare (qq{
			SELECT DISTINCT STORE_CODE
			FROM METRO_IT_INSTOCK_DEPT
			WHERE STORE_FORMAT = 8
			GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0
			ORDER BY 1
			});								 
$st->execute();

while(my $s = $st->fetchrow_hashref()){

	$a += 6, $counter = 0, $loc_pt = $a-2;

	$worksheet->set_column( $col, $col+1, 7 );
	$worksheet->set_column( $col+2, $col+2, 5 );

	foreach my $i ("COUNT", "IN STOCK", "%") {
		$worksheet->write($a-1, $col++, $i, $subhead);
	}

	$col -= 3;

	&query_dept_store( $store = $s->{STORE_CODE} );

	$a = 0, $counter = 0, $stopper = 1, $col += 3;

}
	
$st->finish();

#========== TOTAL WAREHOUSES

$a += 6, $counter = 0;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $headN);
}

$col -= 3;

&query_dept_store( $tot_wh = 1 );

$a = 0, $counter = 0, $col += 3, $tot_wh = 0;

}

sub call_str_merchandise {

$a=6, $counter=0;

$worksheet->write($a-6, 3, "Replenishment In-Stock", $bold);
$worksheet->write($a-5, 3, $day . '-' . $month_to_text . '-' .$year, $bold2);

$worksheet->write($a-3, 3, "Summary", $bold);

&heading_3;

$worksheet->merge_range( $a-2, 3, $a, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a, 6, 'Format', $subhead );

$a += 1;

&query_summary_merchandise($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 1, $matured_flg2 = 1, $summary_label = 'COMP');
&query_summary_merchandise($new_flg1 = 0, $new_flg2 = 0, $matured_flg1 = 0, $matured_flg2 = 0, $summary_label = 'NEW');
&query_summary_merchandise($new_flg1 = 0, $new_flg2 = 1, $matured_flg1 = 0, $matured_flg2 = 1, $summary_label = 'ALL');

$a += 5;

$worksheet->write($a-4, 3, "Per Store", $bold);

&heading_3;

$worksheet->merge_range( $a-2, 3, $a, 3, 'Type', $subhead );
$worksheet->merge_range( $a-2, 4, $a, 4, 'Type', $subhead );
$worksheet->merge_range( $a-2, 5, $a, 5, 'Code', $subhead );
$worksheet->merge_range( $a-2, 6, $a, 6, 'Description', $subhead );

$a += 1;

&query_by_store_merchandise($store_format = 'DEPARTMENT STORE');
&query_by_store_merchandise($store_format = 'SUPERMARKET');
&query_by_store_merchandise($store_format = 'HYPERMARKET');

}

sub call_summary_division {

##=============VISAYAS
$a = 6, $col = 7, $counter = 0; 

&heading_4;

foreach my $i ("Visayas Stores", "Luzon Stores", "All Stores", "All Warehouses") {
	$worksheet->set_column( $col, $col+1, 7 );
	$worksheet->set_column( $col+2, $col+2, 5 );
	
	$worksheet->merge_range( $a-2, $col, $a-2, $col+2, $i, $subhead );
	
	foreach my $i ("COUNT", "IN STOCK", "%") {
		$worksheet->write($a-1, $col++, $i, $subhead);
	}
	
}

&query_summary_division;


}

#create new sheet
sub new_sheet{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(85);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
$worksheet->set_print_scale( 100 );
$worksheet->set_margins( 0.05 );
$worksheet->conditional_formatting( 'F9:V2000',  { type => 'cell', criteria => '<', value => 0, format => $down } );
$worksheet->set_column( 0, 0, undef, undef, 1 );
$worksheet->set_column( 1, 2, 3 );
$worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 23 );

}

#headers
sub heading_orig {

$worksheet->write(0, 2, "Replenishment In-Stock - " . $comment , $bold);
$worksheet->write(1, 2, $day . '-' . $month_to_text . '-' .$year, $bold2);
$worksheet->merge_range( 4, 2, 5, 2, 'Type', $subhead );
$worksheet->merge_range( 4, 3, 5, 3, 'Type', $subhead );
$worksheet->merge_range( 4, 4, 5, 4, 'Type', $subhead );
$worksheet->merge_range( 4, 5, 5, 5, 'Code', $subhead );
$worksheet->merge_range( 4, 6, 5, 6, 'Description', $subhead );

}

sub heading_3 {

$worksheet->merge_range( $a-2, 7, $a-1, 9, 'General Merchandise', $subhead );
$worksheet->merge_range( $a-2, 10, $a-1, 12, 'Supermarket', $subhead );
$worksheet->merge_range( $a-2, 13, $a-1, 15, 'Total', $subhead );

foreach my $i ( 7, 10, 13 ) {
	$worksheet->write($a, $i, "SKU Count", $subhead);
	$worksheet->write($a, $i+1, "In-Stock", $subhead);
	$worksheet->write($a, $i+2, "%", $subhead);
}

}

sub heading_4 {

$worksheet->write(0, 2, "Replenishment In-Stock " . $comment , $bold);
$worksheet->write(1, 2, $day . '-' . $month_to_text . '-' .$year, $bold2);
$worksheet->merge_range( 4, 2, 5, 2, 'Type', $subhead );
$worksheet->merge_range( 4, 3, 5, 3, 'Type', $subhead );
$worksheet->merge_range( 4, 4, 5, 6, 'Description', $subhead );

}

#sheet 1
sub query_summary_merchandise {

$sls = $dbh->prepare (qq{
	SELECT SUM(TOT_ITEMS_DS) TOT_ITEMS_DS, SUM(REP_ITEMS_DS) REP_ITEMS_DS, SUM(TOT_ITEMS_SU) TOT_ITEMS_SU, SUM(REP_ITEMS_SU) REP_ITEMS_SU, SUM(nvl(TOT_ITEMS_DS,0)+nvl(TOT_ITEMS_SU,0)) TOT_ITEMS, SUM(nvl(REP_ITEMS_DS,0)+nvl(REP_ITEMS_SU,0)) REP_ITEMS FROM
		(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_MARGIN_DEPT WHERE STORE_FORMAT <> 3 AND STORE_FORMAT < 5)A LEFT JOIN
		(SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_DS, SUM(REPL_WITH_SOH) AS REP_ITEMS_DS
		FROM METRO_IT_INSTOCK_DEPT
		WHERE MERCH_GROUP_CODE = 'DS' 
			AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
		GROUP BY UPPER(STORE_FORMAT_DESC))B ON A.FORMAT = B.FORMAT
		LEFT JOIN
		(SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_SU, SUM(REPL_WITH_SOH) AS REP_ITEMS_SU
		FROM METRO_IT_INSTOCK_DEPT
		WHERE MERCH_GROUP_CODE = 'SU' 
			AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
			AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
		GROUP BY UPPER(STORE_FORMAT_DESC))C ON A.FORMAT = C.FORMAT	
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
		SELECT A.FORMAT, TOT_ITEMS_DS, REP_ITEMS_DS, TOT_ITEMS_SU, REP_ITEMS_SU, nvl(TOT_ITEMS_DS,0)+nvl(TOT_ITEMS_SU,0) TOT_ITEMS, nvl(REP_ITEMS_DS,0)+nvl(REP_ITEMS_SU,0) REP_ITEMS FROM
			(SELECT DISTINCT UPPER(STORE_FORMAT_DESC) FORMAT FROM METRO_IT_MARGIN_DEPT WHERE STORE_FORMAT <> 3 AND STORE_FORMAT < 5)A LEFT JOIN
			(SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_DS, SUM(REPL_WITH_SOH) AS REP_ITEMS_DS
			FROM METRO_IT_INSTOCK_DEPT
			WHERE MERCH_GROUP_CODE = 'DS' 
				AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY UPPER(STORE_FORMAT_DESC))B ON A.FORMAT = B.FORMAT
			LEFT JOIN
			(SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_SU, SUM(REPL_WITH_SOH) AS REP_ITEMS_SU
			FROM METRO_IT_INSTOCK_DEPT
			WHERE MERCH_GROUP_CODE = 'SU' 
				AND ((NEW_FLG = '$new_flg1' OR NEW_FLG = '$new_flg2') AND (MATURED_FLG = '$matured_flg1' OR MATURED_FLG = '$matured_flg2')) 
				AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY UPPER(STORE_FORMAT_DESC))C ON A.FORMAT = C.FORMAT
		ORDER BY 1
		});
	$sls1->execute();
						
		while(my $s = $sls1->fetchrow_hashref()){
								
		$worksheet->merge_range( $a, 4, $a, 6, $s->{FORMAT}, $desc );
		
		$worksheet->write($a,7, $s->{TOT_ITEMS_DS},$border1);
		$worksheet->write($a,8, $s->{REP_ITEMS_DS},$border1);
		
		if ($s->{TOT_ITEMS_DS} <= 0){
			$worksheet->write($a,9, "",$subt);}
		else{
			$worksheet->write($a,9, '=IF(ISERROR('. $s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS} .'),"",('.$s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS}. '))',$subt);}
			
		$worksheet->write($a,10, $s->{TOT_ITEMS_SU}, $border1);
		$worksheet->write($a,11, $s->{REP_ITEMS_SU},$border1);
		
		if ($s->{TOT_ITEMS_SU} <= 0){
			$worksheet->write($a,12, "",$subt);}
		else{
			$worksheet->write($a,12, '=IF(ISERROR('. $s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU} .'),"",('.$s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU}. '))',$subt);}
		
		$worksheet->write($a,13, $s->{TOT_ITEMS},$border1);
		$worksheet->write($a,14, $s->{REP_ITEMS},$border1);
		
		if ($s->{TOT_ITEMS} <= 0){
			$worksheet->write($a,15, "",$subt);}
		else{
			$worksheet->write($a,15, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$subt);}
												
		$a++;
		$counter++;
					
	}
	
	$worksheet->write($a,7, $s->{TOT_ITEMS_DS},$bodyNum);
	$worksheet->write($a,8, $s->{REP_ITEMS_DS},$bodyNum);
	
	if ($s->{TOT_ITEMS_DS} <= 0){
		$worksheet->write($a,9, "",$bodyPct);}
	else{
		$worksheet->write($a,9, '=IF(ISERROR('. $s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS} .'),"",('.$s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS}. '))',$bodyPct);}
		
	$worksheet->write($a,10, $s->{TOT_ITEMS_SU},$bodyNum);
	$worksheet->write($a,11, $s->{REP_ITEMS_SU},$bodyNum);
	
	if ($s->{TOT_ITEMS_SU} <= 0){
		$worksheet->write($a,12, "",$bodyPct);}
	else{
		$worksheet->write($a,12, '=IF(ISERROR('. $s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU} .'),"",('.$s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU}. '))',$bodyPct);}
	
	$worksheet->write($a,13, $s->{TOT_ITEMS},$bodyNum);
	$worksheet->write($a,14, $s->{REP_ITEMS},$bodyNum);
	
	if ($s->{TOT_ITEMS} <= 0){
		$worksheet->write($a,15, "",$bodyPct);}
	else{
		$worksheet->write($a,15, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$bodyPct);}
		
	$worksheet->merge_range( $a, 4, $a, 6, 'Total', $bodyN );
	$worksheet->merge_range( $a-$counter, 3, $a, 3, $summary_label, $border2 );
	
	$counter = 0;
	$a++;
}

$sls->finish();
$sls1->finish();

}

sub query_by_store_merchandise {

$sls = $dbh->prepare (qq{
	SELECT SUM(TOT_ITEMS_DS) TOT_ITEMS_DS, SUM(REP_ITEMS_DS) REP_ITEMS_DS, SUM(TOT_ITEMS_SU) TOT_ITEMS_SU, SUM(REP_ITEMS_SU) REP_ITEMS_SU, SUM(nvl(TOT_ITEMS_DS,0)+nvl(TOT_ITEMS_SU,0)) TOT_ITEMS, SUM(nvl(REP_ITEMS_DS,0)+nvl(REP_ITEMS_SU,0)) REP_ITEMS 
	FROM
	  (SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_NAME, MATURED_FLG, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_DS, SUM(REPL_WITH_SOH) AS REP_ITEMS_DS, 0 AS TOT_ITEMS_SU, 0 AS REP_ITEMS_SU
	  FROM METRO_IT_INSTOCK_DEPT
	  WHERE MERCH_GROUP_CODE = 'DS' AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
	  GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_NAME, MATURED_FLG
	  UNION ALL
	  SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_NAME, MATURED_FLG, 0 AS TOT_ITEMS_DS, 0 AS REP_ITEMS_DS, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_SU, SUM(REPL_WITH_SOH) AS REP_ITEMS_SU
	  FROM METRO_IT_INSTOCK_DEPT
	  WHERE MERCH_GROUP_CODE = 'SU' AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
	  GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_NAME, MATURED_FLG)
	});
$sls->execute();
					
	while(my $s = $sls->fetchrow_hashref()){

	$sls1 = $dbh->prepare (qq{
		SELECT MATURED_FLG, CASE WHEN MATURED_FLG = 1 THEN 'COMP' ELSE 'NEW' END AS FLG_DESC, SUM(TOT_ITEMS_DS) TOT_ITEMS_DS, SUM(REP_ITEMS_DS) REP_ITEMS_DS, SUM(TOT_ITEMS_SU) TOT_ITEMS_SU, SUM(REP_ITEMS_SU) REP_ITEMS_SU, SUM(nvl(TOT_ITEMS_DS,0)+nvl(TOT_ITEMS_SU,0)) TOT_ITEMS, SUM(nvl(REP_ITEMS_DS,0)+nvl(REP_ITEMS_SU,0)) REP_ITEMS 
		FROM
		  (SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_NAME, MATURED_FLG, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_DS, SUM(REPL_WITH_SOH) AS REP_ITEMS_DS, 0 AS TOT_ITEMS_SU, 0 AS REP_ITEMS_SU
		  FROM METRO_IT_INSTOCK_DEPT
		  WHERE MERCH_GROUP_CODE = 'DS' AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
		  GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_NAME, MATURED_FLG
		  UNION ALL
		  SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_NAME, MATURED_FLG, 0 AS TOT_ITEMS_DS, 0 AS REP_ITEMS_DS, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_SU, SUM(REPL_WITH_SOH) AS REP_ITEMS_SU
		  FROM METRO_IT_INSTOCK_DEPT
		  WHERE MERCH_GROUP_CODE = 'SU' AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
		  GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_NAME, MATURED_FLG)
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
			SELECT STORE_CODE, STORE_NAME STORE_DESCRIPTION, SUM(TOT_ITEMS_DS) TOT_ITEMS_DS, SUM(REP_ITEMS_DS) REP_ITEMS_DS, SUM(TOT_ITEMS_SU) TOT_ITEMS_SU, SUM(REP_ITEMS_SU) REP_ITEMS_SU, SUM(nvl(TOT_ITEMS_DS,0)+nvl(TOT_ITEMS_SU,0)) TOT_ITEMS, SUM(nvl(REP_ITEMS_DS,0)+nvl(REP_ITEMS_SU,0)) REP_ITEMS 
			FROM
			  (SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_NAME, MATURED_FLG, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_DS, SUM(REPL_WITH_SOH) AS REP_ITEMS_DS, 0 AS TOT_ITEMS_SU, 0 AS REP_ITEMS_SU
			  FROM METRO_IT_INSTOCK_DEPT
			  WHERE MERCH_GROUP_CODE = 'DS' AND MATURED_FLG = '$flg' AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			  GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_NAME, MATURED_FLG
			  UNION ALL
			  SELECT UPPER(STORE_FORMAT_DESC) AS FORMAT, STORE_CODE, STORE_NAME, MATURED_FLG, 0 AS TOT_ITEMS_DS, 0 AS REP_ITEMS_DS, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_SU, SUM(REPL_WITH_SOH) AS REP_ITEMS_SU
			  FROM METRO_IT_INSTOCK_DEPT
			  WHERE MERCH_GROUP_CODE = 'SU' AND MATURED_FLG = '$flg' AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			  GROUP BY UPPER(STORE_FORMAT_DESC), STORE_CODE, STORE_NAME, MATURED_FLG)
			  WHERE FORMAT = '$store_format'
			GROUP BY STORE_CODE, STORE_NAME
			ORDER BY 1
			});
		$sls2->execute();
			
			while(my $s = $sls2->fetchrow_hashref()){
									
			$worksheet->write( $a, 5, $s->{STORE_CODE}, $desc );
			$worksheet->write( $a, 6, $s->{STORE_DESCRIPTION}, $desc );
			
			$worksheet->write($a,7, $s->{TOT_ITEMS_DS},$border1);
			$worksheet->write($a,8, $s->{REP_ITEMS_DS},$border1);
		
			if ($s->{TOT_ITEMS_DS} <= 0){
				$worksheet->write($a,9, "",$subt);}
			else{
				$worksheet->write($a,9, '=IF(ISERROR('. $s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS} .'),"",('.$s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS}. '))',$subt);}
				
			$worksheet->write($a,10, $s->{TOT_ITEMS_SU}, $border1);
			$worksheet->write($a,11, $s->{REP_ITEMS_SU},$border1);
			
			if ($s->{TOT_ITEMS_SU} <= 0){
				$worksheet->write($a,12, "",$subt);}
			else{
				$worksheet->write($a,12, '=IF(ISERROR('. $s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU} .'),"",('.$s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU}. '))',$subt);}
			
			$worksheet->write($a,13, $s->{TOT_ITEMS},$border1);
			$worksheet->write($a,14, $s->{REP_ITEMS},$border1);
			
			if ($s->{TOT_ITEMS} <= 0){
				$worksheet->write($a,15, "",$subt);}
			else{
				$worksheet->write($a,15, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$subt);}
													
			$a++;
			$counter++;
						
		}
		
		$worksheet->write($a,7, $s->{TOT_ITEMS_DS},$bodyNum);
		$worksheet->write($a,8, $s->{REP_ITEMS_DS},$bodyNum);
		
		if ($s->{TOT_ITEMS_DS} <= 0){
			$worksheet->write($a,9, "",$bodyPct);}
		else{
			$worksheet->write($a,9, '=IF(ISERROR('. $s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS} .'),"",('.$s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS}. '))',$bodyPct);}
			
		$worksheet->write($a,10, $s->{TOT_ITEMS_SU},$bodyNum);
		$worksheet->write($a,11, $s->{REP_ITEMS_SU},$bodyNum);
		
		if ($s->{TOT_ITEMS_SU} <= 0){
			$worksheet->write($a,12, "",$bodyPct);}
		else{
			$worksheet->write($a,12, '=IF(ISERROR('. $s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU} .'),"",('.$s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU}. '))',$bodyPct);}
		
		$worksheet->write($a,13, $s->{TOT_ITEMS},$bodyNum);
		$worksheet->write($a,14, $s->{REP_ITEMS},$bodyNum);
		
		if ($s->{TOT_ITEMS} <= 0){
			$worksheet->write($a,15, "",$bodyPct);}
		else{
			$worksheet->write($a,15, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$bodyPct);}
		
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $flg_desc, $border2 );
		
		$a++;
		$counter = 0;
	}
	
	$worksheet->write($a,7, $s->{TOT_ITEMS_DS},$bodyNum);
	$worksheet->write($a,8, $s->{REP_ITEMS_DS},$bodyNum);
	
	if ($s->{TOT_ITEMS_DS} <= 0){
		$worksheet->write($a,9, "",$bodyPct);}
	else{
		$worksheet->write($a,9, '=IF(ISERROR('. $s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS} .'),"",('.$s->{REP_ITEMS_DS}/$s->{TOT_ITEMS_DS}. '))',$bodyPct);}
		
	$worksheet->write($a,10, $s->{TOT_ITEMS_SU},$bodyNum);
	$worksheet->write($a,11, $s->{REP_ITEMS_SU},$bodyNum);
	
	if ($s->{TOT_ITEMS_SU} <= 0){
		$worksheet->write($a,12, "",$bodyPct);}
	else{
		$worksheet->write($a,12, '=IF(ISERROR('. $s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU} .'),"",('.$s->{REP_ITEMS_SU}/$s->{TOT_ITEMS_SU}. '))',$bodyPct);}
	
	$worksheet->write($a,13, $s->{TOT_ITEMS},$bodyNum);
	$worksheet->write($a,14, $s->{REP_ITEMS},$bodyNum);
	
	if ($s->{TOT_ITEMS} <= 0){
		$worksheet->write($a,15, "",$bodyPct);}
	else{
		$worksheet->write($a,15, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$bodyPct);}
	
	$worksheet->merge_range( $a, 4, $a, 6, 'Total ' . $store_format, $bodyN );
	$worksheet->merge_range( $format_counter, 3, $a, 3, $store_format, $border2 );
	
	$a++;
}

$sls->finish();
$sls1->finish();
$sls2->finish();

}

#sheet 2
sub query_summary_division {
		
$sls = $dbh->prepare (qq{
			SELECT 
			  SUM(TOT_ITEMS_VIS) AS TOT_ITEMS_VIS, SUM(REP_ITEMS_VIS) AS REP_ITEMS_VIS, 
			  SUM(TOT_ITEMS_LUZ) AS TOT_ITEMS_LUZ, SUM(REP_ITEMS_LUZ) AS REP_ITEMS_LUZ, 
			  SUM(NVL(TOT_ITEMS_VIS,0)+NVL(TOT_ITEMS_LUZ,0)) AS TOT_ITEMS, SUM(NVL(REP_ITEMS_VIS,0)+NVL(REP_ITEMS_LUZ,0)) AS REP_ITEMS, 
			  SUM(TOT_ITEMS_WH) AS TOT_ITEMS_WH, SUM(REP_ITEMS_WH) AS REP_ITEMS_WH
			FROM
			(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
				SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_VIS, SUM(REPL_WITH_SOH) AS REP_ITEMS_VIS, 
				0 AS TOT_ITEMS_LUZ, 0 AS REP_ITEMS_LUZ,
				0 AS TOT_ITEMS_WH, 0 AS REP_ITEMS_WH
			FROM METRO_IT_INSTOCK_DEPT
			WHERE UPPER(REGION) LIKE '%VISAYAS%' AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC
			UNION ALL
			SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
				0 AS TOT_ITEMS_VIS, 0 AS REP_ITEMS_VIS, 
				SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_LUZ, SUM(REPL_WITH_SOH) AS REP_ITEMS_LUZ,
				0 AS TOT_ITEMS_WH, 0 AS REP_ITEMS_WH
			FROM METRO_IT_INSTOCK_DEPT
			WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY') AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
			GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC
			UNION ALL
			SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
				0 AS TOT_ITEMS_VIS, 0 AS REP_ITEMS_VIS, 
				0 AS TOT_ITEMS_LUZ, 0 AS REP_ITEMS_LUZ,
				SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_WH, SUM(REPL_WITH_SOH) AS REP_ITEMS_WH
			FROM METRO_IT_INSTOCK_DEPT
			WHERE STORE_FORMAT = 8
			GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC)
			});								 
$sls->execute();

	while(my $s = $sls->fetchrow_hashref()){
				
	$sls1 = $dbh->prepare (qq{
			SELECT * FROM (
				SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, 
				  SUM(TOT_ITEMS_VIS) AS TOT_ITEMS_VIS, SUM(REP_ITEMS_VIS) AS REP_ITEMS_VIS, 
				  SUM(TOT_ITEMS_LUZ) AS TOT_ITEMS_LUZ, SUM(REP_ITEMS_LUZ) AS REP_ITEMS_LUZ, 
				  SUM(NVL(TOT_ITEMS_VIS,0)+NVL(TOT_ITEMS_LUZ,0)) AS TOT_ITEMS, SUM(NVL(REP_ITEMS_VIS,0)+NVL(REP_ITEMS_LUZ,0)) AS REP_ITEMS, 
				  SUM(TOT_ITEMS_WH) AS TOT_ITEMS_WH, SUM(REP_ITEMS_WH) AS REP_ITEMS_WH,
				  SUM(NVL(TOT_ITEMS_VIS,0)+NVL(TOT_ITEMS_LUZ,0)+NVL(TOT_ITEMS_WH,0)) AS TOT
				FROM
					(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
						SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_VIS, SUM(REPL_WITH_SOH) AS REP_ITEMS_VIS, 
						0 AS TOT_ITEMS_LUZ, 0 AS REP_ITEMS_LUZ,
						0 AS TOT_ITEMS_WH, 0 AS REP_ITEMS_WH
					FROM METRO_IT_INSTOCK_DEPT
					WHERE UPPER(REGION) LIKE '%VISAYAS%' AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC
					UNION ALL
					SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
						0 AS TOT_ITEMS_VIS, 0 AS REP_ITEMS_VIS, 
						SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_LUZ, SUM(REPL_WITH_SOH) AS REP_ITEMS_LUZ,
						0 AS TOT_ITEMS_WH, 0 AS REP_ITEMS_WH
					FROM METRO_IT_INSTOCK_DEPT
					WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY') AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC
					UNION ALL
					SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
						0 AS TOT_ITEMS_VIS, 0 AS REP_ITEMS_VIS, 
						0 AS TOT_ITEMS_LUZ, 0 AS REP_ITEMS_LUZ,
						SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_WH, SUM(REPL_WITH_SOH) AS REP_ITEMS_WH
					FROM METRO_IT_INSTOCK_DEPT
					WHERE STORE_FORMAT = 8
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC)
				GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC)
			WHERE TOT <> 0 ORDER BY 1
				});								 
	$sls1->execute();

		while(my $s = $sls1->fetchrow_hashref()){
			$merch_group_code = $s->{MERCH_GROUP_CODE};
			$merch_group_desc = $s->{MERCH_GROUP_DESC}; 
			
			$sls2 = $dbh->prepare (qq{
				SELECT * FROM (
					SELECT GROUP_CODE, GROUP_DESC, 
					  SUM(TOT_ITEMS_VIS) AS TOT_ITEMS_VIS, SUM(REP_ITEMS_VIS) AS REP_ITEMS_VIS, 
					  SUM(TOT_ITEMS_LUZ) AS TOT_ITEMS_LUZ, SUM(REP_ITEMS_LUZ) AS REP_ITEMS_LUZ, 
					  SUM(NVL(TOT_ITEMS_VIS,0)+NVL(TOT_ITEMS_LUZ,0)) AS TOT_ITEMS, SUM(NVL(REP_ITEMS_VIS,0)+NVL(REP_ITEMS_LUZ,0)) AS REP_ITEMS, 
					  SUM(TOT_ITEMS_WH) AS TOT_ITEMS_WH, SUM(REP_ITEMS_WH) AS REP_ITEMS_WH,
					  SUM(NVL(TOT_ITEMS_VIS,0)+NVL(TOT_ITEMS_LUZ,0)+NVL(TOT_ITEMS_WH,0)) AS TOT
					FROM
						(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
							SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_VIS, SUM(REPL_WITH_SOH) AS REP_ITEMS_VIS, 
							0 AS TOT_ITEMS_LUZ, 0 AS REP_ITEMS_LUZ,
							0 AS TOT_ITEMS_WH, 0 AS REP_ITEMS_WH
						FROM METRO_IT_INSTOCK_DEPT
						WHERE UPPER(REGION) LIKE '%VISAYAS%' 
							AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
							AND MERCH_GROUP_CODE = '$merch_group_code' 
						GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC
						UNION ALL
						SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
							0 AS TOT_ITEMS_VIS, 0 AS REP_ITEMS_VIS, 
							SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_LUZ, SUM(REPL_WITH_SOH) AS REP_ITEMS_LUZ,
							0 AS TOT_ITEMS_WH, 0 AS REP_ITEMS_WH
						FROM METRO_IT_INSTOCK_DEPT
						WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY') 
							AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
							AND MERCH_GROUP_CODE = '$merch_group_code' 
						GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC
						UNION ALL
						SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
							0 AS TOT_ITEMS_VIS, 0 AS REP_ITEMS_VIS, 
							0 AS TOT_ITEMS_LUZ, 0 AS REP_ITEMS_LUZ,
							SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_WH, SUM(REPL_WITH_SOH) AS REP_ITEMS_WH
						FROM METRO_IT_INSTOCK_DEPT
						WHERE STORE_FORMAT = 8 
							AND MERCH_GROUP_CODE = '$merch_group_code' 
						GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC)
					GROUP BY GROUP_CODE, GROUP_DESC)
				WHERE TOT <> 0 ORDER BY 1
					});	
			$sls2->execute();
			
			$mgc_counter = $a;
			while(my $s = $sls2->fetchrow_hashref()){
				$group_code = $s->{GROUP_CODE};
				$group_desc = $s->{GROUP_DESC};
						
				$sls3 = $dbh->prepare (qq{
					SELECT * FROM (
						SELECT DIVISION, DIVISION_DESC, 
						  SUM(TOT_ITEMS_VIS) AS TOT_ITEMS_VIS, SUM(REP_ITEMS_VIS) AS REP_ITEMS_VIS, 
						  SUM(TOT_ITEMS_LUZ) AS TOT_ITEMS_LUZ, SUM(REP_ITEMS_LUZ) AS REP_ITEMS_LUZ, 
						  SUM(NVL(TOT_ITEMS_VIS,0)+NVL(TOT_ITEMS_LUZ,0)) AS TOT_ITEMS, SUM(NVL(REP_ITEMS_VIS,0)+NVL(REP_ITEMS_LUZ,0)) AS REP_ITEMS, 
						  SUM(TOT_ITEMS_WH) AS TOT_ITEMS_WH, SUM(REP_ITEMS_WH) AS REP_ITEMS_WH,
						  SUM(NVL(TOT_ITEMS_VIS,0)+NVL(TOT_ITEMS_LUZ,0)+NVL(TOT_ITEMS_WH,0)) AS TOT
						FROM
							(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
								SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_VIS, SUM(REPL_WITH_SOH) AS REP_ITEMS_VIS, 
								0 AS TOT_ITEMS_LUZ, 0 AS REP_ITEMS_LUZ,
								0 AS TOT_ITEMS_WH, 0 AS REP_ITEMS_WH
							FROM METRO_IT_INSTOCK_DEPT
							WHERE UPPER(REGION) LIKE '%VISAYAS%' 
								AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
								AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
							GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC
							UNION ALL
							SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
								0 AS TOT_ITEMS_VIS, 0 AS REP_ITEMS_VIS, 
								SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_LUZ, SUM(REPL_WITH_SOH) AS REP_ITEMS_LUZ,
								0 AS TOT_ITEMS_WH, 0 AS REP_ITEMS_WH
							FROM METRO_IT_INSTOCK_DEPT
							WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY') 
								AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
								AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
							GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC
							UNION ALL
							SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC,
								0 AS TOT_ITEMS_VIS, 0 AS REP_ITEMS_VIS, 
								0 AS TOT_ITEMS_LUZ, 0 AS REP_ITEMS_LUZ,
								SUM(TOT_REPL_ITEMS) AS TOT_ITEMS_WH, SUM(REPL_WITH_SOH) AS REP_ITEMS_WH
							FROM METRO_IT_INSTOCK_DEPT
							WHERE STORE_FORMAT = 8 
								AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code' 
							GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC, GROUP_CODE, GROUP_DESC, DIVISION, DIVISION_DESC)
						GROUP BY DIVISION, DIVISION_DESC)
					WHERE TOT <> 0 ORDER BY 1
					});
				$sls3->execute();
				
				$grp_counter = $a;
				while(my $s = $sls3->fetchrow_hashref()){
					
					$worksheet->merge_range( $a, 4, $a, 6, $s->{DIVISION_DESC}, $desc );
					
					$worksheet->write($a,7, $s->{TOT_ITEMS_VIS},$border1);
					$worksheet->write($a,8, $s->{REP_ITEMS_VIS},$border1);
					
					if ($s->{TOT_ITEMS_VIS} <= 0){
						$worksheet->write($a,9, "",$subt);}
					else{
						$worksheet->write($a,9, '=IF(ISERROR('. $s->{REP_ITEMS_VIS}/$s->{TOT_ITEMS_VIS} .'),"",('.$s->{REP_ITEMS_VIS}/$s->{TOT_ITEMS_VIS}. '))',$subt);}
						
					$worksheet->write($a,10, $s->{TOT_ITEMS_LUZ}, $border1);
					$worksheet->write($a,11, $s->{REP_ITEMS_LUZ},$border1);
					
					if ($s->{TOT_ITEMS_LUZ} <= 0){
						$worksheet->write($a,12, "",$subt);}
					else{
						$worksheet->write($a,12, '=IF(ISERROR('. $s->{REP_ITEMS_LUZ}/$s->{TOT_ITEMS_LUZ} .'),"",('.$s->{REP_ITEMS_LUZ}/$s->{TOT_ITEMS_LUZ}. '))',$subt);}
					
					$worksheet->write($a,13, $s->{TOT_ITEMS},$border1);
					$worksheet->write($a,14, $s->{REP_ITEMS},$border1);
					
					if ($s->{TOT_ITEMS} <= 0){
						$worksheet->write($a,15, "",$subt);}
					else{
						$worksheet->write($a,15, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$subt);}
					
					$worksheet->write($a,16, $s->{TOT_ITEMS_WH},$border1);
					$worksheet->write($a,17, $s->{REP_ITEMS_WH},$border1);
					
					if ($s->{TOT_ITEMS_WH} <= 0){
						$worksheet->write($a,18, "",$subt);}
					else{
						$worksheet->write($a,18, '=IF(ISERROR('. $s->{REP_ITEMS_WH}/$s->{TOT_ITEMS_WH} .'),"",('.$s->{REP_ITEMS_WH}/$s->{TOT_ITEMS_WH}. '))',$subt);}
					
					$counter = 0;
					$a++;
				}
				
				$worksheet->write($a,7, $s->{TOT_ITEMS_VIS},$bodyNum);
				$worksheet->write($a,8, $s->{REP_ITEMS_VIS},$bodyNum);
					
				if ($s->{TOT_ITEMS_VIS} <= 0){
					$worksheet->write($a,9, "",$bodyPct);}
				else{
					$worksheet->write($a,9, '=IF(ISERROR('. $s->{REP_ITEMS_VIS}/$s->{TOT_ITEMS_VIS} .'),"",('.$s->{REP_ITEMS_VIS}/$s->{TOT_ITEMS_VIS}. '))',$bodyPct);}
					
				$worksheet->write($a,10, $s->{TOT_ITEMS_LUZ}, $bodyNum);
				$worksheet->write($a,11, $s->{REP_ITEMS_LUZ},$bodyNum);
				
				if ($s->{TOT_ITEMS_LUZ} <= 0){
					$worksheet->write($a,12, "",$bodyPct);}
				else{
					$worksheet->write($a,12, '=IF(ISERROR('. $s->{REP_ITEMS_LUZ}/$s->{TOT_ITEMS_LUZ} .'),"",('.$s->{REP_ITEMS_LUZ}/$s->{TOT_ITEMS_LUZ}. '))',$bodyPct);}
				
				$worksheet->write($a,13, $s->{TOT_ITEMS},$bodyNum);
				$worksheet->write($a,14, $s->{REP_ITEMS},$bodyNum);
				
				if ($s->{TOT_ITEMS} <= 0){
					$worksheet->write($a,15, "",$bodyPct);}
				else{
					$worksheet->write($a,15, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$bodyPct);}
				
				$worksheet->write($a,16, $s->{TOT_ITEMS_WH},$bodyNum);
				$worksheet->write($a,17, $s->{REP_ITEMS_WH},$bodyNum);
				
				if ($s->{TOT_ITEMS_WH} <= 0){
					$worksheet->write($a,18, "",$bodyPct);}
				else{
					$worksheet->write($a,18, '=IF(ISERROR('. $s->{REP_ITEMS_WH}/$s->{TOT_ITEMS_WH} .'),"",('.$s->{REP_ITEMS_WH}/$s->{TOT_ITEMS_WH}. '))',$bodyPct);}

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
			
			$worksheet->write($a,7, $s->{TOT_ITEMS_VIS},$headNumber);
			$worksheet->write($a,8, $s->{REP_ITEMS_VIS},$headNumber);
				
			if ($s->{TOT_ITEMS_VIS} <= 0){
				$worksheet->write($a,9, "",$headPct);}
			else{
				$worksheet->write($a,9, '=IF(ISERROR('. $s->{REP_ITEMS_VIS}/$s->{TOT_ITEMS_VIS} .'),"",('.$s->{REP_ITEMS_VIS}/$s->{TOT_ITEMS_VIS}. '))',$headPct);}
				
			$worksheet->write($a,10, $s->{TOT_ITEMS_LUZ}, $headNumber);
			$worksheet->write($a,11, $s->{REP_ITEMS_LUZ},$headNumber);
			
			if ($s->{TOT_ITEMS_LUZ} <= 0){
				$worksheet->write($a,12, "",$headPct);}
			else{
				$worksheet->write($a,12, '=IF(ISERROR('. $s->{REP_ITEMS_LUZ}/$s->{TOT_ITEMS_LUZ} .'),"",('.$s->{REP_ITEMS_LUZ}/$s->{TOT_ITEMS_LUZ}. '))',$headPct);}
			
			$worksheet->write($a,13, $s->{TOT_ITEMS},$headNumber);
			$worksheet->write($a,14, $s->{REP_ITEMS},$headNumber);
			
			if ($s->{TOT_ITEMS} <= 0){
				$worksheet->write($a,15, "",$headPct);}
			else{
				$worksheet->write($a,15, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$headPct);}
			
			$worksheet->write($a,16, $s->{TOT_ITEMS_WH},$headNumber);
			$worksheet->write($a,17, $s->{REP_ITEMS_WH},$headNumber);
			
			if ($s->{TOT_ITEMS_WH} <= 0){
				$worksheet->write($a,18, "",$headPct);}
			else{
				$worksheet->write($a,18, '=IF(ISERROR('. $s->{REP_ITEMS_WH}/$s->{TOT_ITEMS_WH} .'),"",('.$s->{REP_ITEMS_WH}/$s->{TOT_ITEMS_WH}. '))',$headPct);}
			
			$a++;
		}

	$worksheet->write($a,7, $s->{TOT_ITEMS_VIS},$headNumber);
	$worksheet->write($a,8, $s->{REP_ITEMS_VIS},$headNumber);
		
	if ($s->{TOT_ITEMS_VIS} <= 0){
		$worksheet->write($a,9, "",$headPct);}
	else{
		$worksheet->write($a,9, '=IF(ISERROR('. $s->{REP_ITEMS_VIS}/$s->{TOT_ITEMS_VIS} .'),"",('.$s->{REP_ITEMS_VIS}/$s->{TOT_ITEMS_VIS}. '))',$headPct);}
		
	$worksheet->write($a,10, $s->{TOT_ITEMS_LUZ}, $headNumber);
	$worksheet->write($a,11, $s->{REP_ITEMS_LUZ},$headNumber);
	
	if ($s->{TOT_ITEMS_LUZ} <= 0){
		$worksheet->write($a,12, "",$headPct);}
	else{
		$worksheet->write($a,12, '=IF(ISERROR('. $s->{REP_ITEMS_LUZ}/$s->{TOT_ITEMS_LUZ} .'),"",('.$s->{REP_ITEMS_LUZ}/$s->{TOT_ITEMS_LUZ}. '))',$headPct);}
			
	$worksheet->write($a,13, $s->{TOT_ITEMS},$headNumber);
	$worksheet->write($a,14, $s->{REP_ITEMS},$headNumber);
		
	if ($s->{TOT_ITEMS} <= 0){
		$worksheet->write($a,15, "",$headPct);}
	else{
		$worksheet->write($a,15, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$headPct);}
	
	$worksheet->write($a,16, $s->{TOT_ITEMS_WH},$headNumber);
	$worksheet->write($a,17, $s->{REP_ITEMS_WH},$headNumber);
	
	if ($s->{TOT_ITEMS_WH} <= 0){
		$worksheet->write($a,18, "",$headPct);}
	else{
		$worksheet->write($a,18, '=IF(ISERROR('. $s->{REP_ITEMS_WH}/$s->{TOT_ITEMS_WH} .'),"",('.$s->{REP_ITEMS_WH}/$s->{TOT_ITEMS_WH}. '))',$headPct);}

}

$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );

$sls->finish();
$sls1->finish();
$sls2->finish();
$sls3->finish();

$counter = 0;

}

#sheet3
sub query_dept_store {

if ($tot_vis eq 1){
	$sls = $dbh->prepare (qq{
			SELECT SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
			FROM METRO_IT_INSTOCK_DEPT
			WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
				FROM METRO_IT_INSTOCK_DEPT
				WHERE UPPER(REGION) LIKE '%VISAYAS%' AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
				GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 )
			});		
}
elsif ($tot_luz eq 1){
	$sls = $dbh->prepare (qq{
			SELECT SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
			FROM METRO_IT_INSTOCK_DEPT
			WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
				FROM METRO_IT_INSTOCK_DEPT
				WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY') 
					AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
				GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 )
			});		
}
elsif ($tot_wh eq 1){
	$sls = $dbh->prepare (qq{
			SELECT SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
			FROM METRO_IT_INSTOCK_DEPT
			WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
				FROM METRO_IT_INSTOCK_DEPT
				WHERE STORE_FORMAT = 8
				GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 )
			});		
}
elsif ($tot eq 1){
	$sls = $dbh->prepare (qq{
			SELECT SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
			FROM METRO_IT_INSTOCK_DEPT
			WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
				FROM METRO_IT_INSTOCK_DEPT
				WHERE UPPER(REGION) <> 'DUMMY'
					AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
				GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 )
			});		
}
else{
	$sls = $dbh->prepare (qq{
			SELECT SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
			FROM METRO_IT_INSTOCK_DEPT
			WHERE STORE_CODE IN ('$store')
			});	
}							 
	
	$sls->execute();
	while(my $s = $sls->fetchrow_hashref()){
	
	if ($tot_vis eq 1){
		$sls1 = $dbh->prepare (qq{
				SELECT A.MERCH_GROUP_CODE, A.MERCH_GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
					(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC
					FROM METRO_IT_INSTOCK_DEPT
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
					LEFT JOIN
					(SELECT MERCH_GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
					FROM METRO_IT_INSTOCK_DEPT
					WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
						FROM METRO_IT_INSTOCK_DEPT
						WHERE UPPER(REGION) LIKE '%VISAYAS%' AND (  (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
						GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 )
					GROUP BY MERCH_GROUP_CODE)B
					ON A.MERCH_GROUP_CODE = B.MERCH_GROUP_CODE ORDER BY 3
				});				
	}
	elsif ($tot_luz eq 1){
		$sls1 = $dbh->prepare (qq{
				SELECT A.MERCH_GROUP_CODE, A.MERCH_GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
					(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC
					FROM METRO_IT_INSTOCK_DEPT
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
					LEFT JOIN
					(SELECT MERCH_GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
					FROM METRO_IT_INSTOCK_DEPT
					WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
						FROM METRO_IT_INSTOCK_DEPT
						WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY') 
							AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
						GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 )
					GROUP BY MERCH_GROUP_CODE)B
					ON A.MERCH_GROUP_CODE = B.MERCH_GROUP_CODE ORDER BY 3
				});				
	}
	elsif ($tot_wh eq 1){
		$sls1 = $dbh->prepare (qq{
				SELECT A.MERCH_GROUP_CODE, A.MERCH_GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
					(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC
					FROM METRO_IT_INSTOCK_DEPT
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
					LEFT JOIN
					(SELECT MERCH_GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
					FROM METRO_IT_INSTOCK_DEPT
					WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
						FROM METRO_IT_INSTOCK_DEPT
						WHERE STORE_FORMAT = 8
						GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 )
					GROUP BY MERCH_GROUP_CODE)B
					ON A.MERCH_GROUP_CODE = B.MERCH_GROUP_CODE ORDER BY 3
				});				
	}
	elsif ($tot eq 1){
		$sls1 = $dbh->prepare (qq{
				SELECT A.MERCH_GROUP_CODE, A.MERCH_GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
					(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC
					FROM METRO_IT_INSTOCK_DEPT
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
					LEFT JOIN
					(SELECT MERCH_GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
					FROM METRO_IT_INSTOCK_DEPT
					WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
						FROM METRO_IT_INSTOCK_DEPT
						WHERE UPPER(REGION) <> 'DUMMY'
							AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
						GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 )
					GROUP BY MERCH_GROUP_CODE)B
					ON A.MERCH_GROUP_CODE = B.MERCH_GROUP_CODE ORDER BY 3
				});				
	}
	else{
		$sls1 = $dbh->prepare (qq{
				SELECT B.STORE_CODE, B.STORE_DESCRIPTION, A.MERCH_GROUP_CODE, A.MERCH_GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
					(SELECT MERCH_GROUP_CODE, MERCH_GROUP_DESC
					FROM METRO_IT_INSTOCK_DEPT
					GROUP BY MERCH_GROUP_CODE, MERCH_GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
					LEFT JOIN
					(SELECT STORE_CODE, STORE_NAME STORE_DESCRIPTION, MERCH_GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
					FROM METRO_IT_INSTOCK_DEPT
					WHERE STORE_CODE IN ('$store')
					GROUP BY STORE_CODE, STORE_NAME, MERCH_GROUP_CODE)B
					ON A.MERCH_GROUP_CODE = B.MERCH_GROUP_CODE ORDER BY 3
				});	
	}							 
		
		$sls1->execute();
		while(my $s = $sls1->fetchrow_hashref()){
			$merch_group_code = $s->{MERCH_GROUP_CODE};
			$merch_group_desc = $s->{MERCH_GROUP_DESC}; 
			
			if ($tot_vis eq 1){
				$loc_code = 'TOTAL';
				$loc_desc = 'VISAYAS';
				$sls2 = $dbh->prepare (qq{
					SELECT A.GROUP_CODE, A.GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
						(SELECT GROUP_CODE, GROUP_DESC
						FROM METRO_IT_INSTOCK_DEPT
						WHERE MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE, GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
						LEFT JOIN
						(SELECT GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
						FROM METRO_IT_INSTOCK_DEPT
						WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
								FROM METRO_IT_INSTOCK_DEPT
								WHERE UPPER(REGION) LIKE '%VISAYAS%' AND (  (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '699A9') )
								GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
							AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE)B
						ON A.GROUP_CODE = B.GROUP_CODE ORDER BY 1
					});	
			}
			elsif ($tot_luz eq 1){
				$loc_code = 'TOTAL';
				$loc_desc = 'LUZON';
				$sls2 = $dbh->prepare (qq{
					SELECT A.GROUP_CODE, A.GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
						(SELECT GROUP_CODE, GROUP_DESC
						FROM METRO_IT_INSTOCK_DEPT
						WHERE MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE, GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
						LEFT JOIN
						(SELECT GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
						FROM METRO_IT_INSTOCK_DEPT
						WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
								FROM METRO_IT_INSTOCK_DEPT
								WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY') 
									AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '699A9') )
								GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
							AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE)B
						ON A.GROUP_CODE = B.GROUP_CODE ORDER BY 1
					});	
			}
			elsif ($tot_wh eq 1){
				$loc_code = 'TOTAL';
				$loc_desc = 'WAREHOUSES';
				$sls2 = $dbh->prepare (qq{
					SELECT A.GROUP_CODE, A.GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
						(SELECT GROUP_CODE, GROUP_DESC
						FROM METRO_IT_INSTOCK_DEPT
						WHERE MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE, GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
						LEFT JOIN
						(SELECT GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
						FROM METRO_IT_INSTOCK_DEPT
						WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
								FROM METRO_IT_INSTOCK_DEPT
								WHERE STORE_FORMAT = 8
								GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
							AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE)B
						ON A.GROUP_CODE = B.GROUP_CODE ORDER BY 1
					});	
			}
			elsif ($tot eq 1){
				$loc_code = 'TOTAL';
				$loc_desc = 'STORES';
				$sls2 = $dbh->prepare (qq{
					SELECT A.GROUP_CODE, A.GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
						(SELECT GROUP_CODE, GROUP_DESC
						FROM METRO_IT_INSTOCK_DEPT
						WHERE MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE, GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
						LEFT JOIN
						(SELECT GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
						FROM METRO_IT_INSTOCK_DEPT
						WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
								FROM METRO_IT_INSTOCK_DEPT
								WHERE UPPER(REGION) <> 'DUMMY'
									AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '699A9') )
								GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
							AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE)B
						ON A.GROUP_CODE = B.GROUP_CODE ORDER BY 1
					});	
			}
			else{
				$loc_code = $s->{STORE_CODE};
				$loc_desc = $s->{STORE_DESCRIPTION};
			
				$sls2 = $dbh->prepare (qq{
					SELECT A.GROUP_CODE, A.GROUP_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
						(SELECT GROUP_CODE, GROUP_DESC
						FROM METRO_IT_INSTOCK_DEPT
						WHERE MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE, GROUP_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
						LEFT JOIN
						(SELECT GROUP_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
						FROM METRO_IT_INSTOCK_DEPT
						WHERE STORE_CODE IN ('$store') AND MERCH_GROUP_CODE = '$merch_group_code'
						GROUP BY GROUP_CODE)B
						ON A.GROUP_CODE = B.GROUP_CODE ORDER BY 1
					});	
			}
			
			$sls2->execute();			
			$mgc_counter = $a;
			while(my $s = $sls2->fetchrow_hashref()){
				$group_code = $s->{GROUP_CODE};
				$group_desc = $s->{GROUP_DESC};
				
				if ($tot_vis eq 1){
					$sls3 = $dbh->prepare (qq{
						SELECT A.DIVISION, A.DIVISION_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
							(SELECT DIVISION, DIVISION_DESC
							FROM METRO_IT_INSTOCK_DEPT
							WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION, DIVISION_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
							LEFT JOIN
							(SELECT DIVISION, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
							FROM METRO_IT_INSTOCK_DEPT
							WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
									FROM METRO_IT_INSTOCK_DEPT
									WHERE UPPER(REGION) LIKE '%VISAYAS%' AND (  (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
									GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
								AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION)B
							ON A.DIVISION = B.DIVISION ORDER BY 1
						});
				}
				elsif ($tot_luz eq 1){
					$sls3 = $dbh->prepare (qq{
						SELECT A.DIVISION, A.DIVISION_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
							(SELECT DIVISION, DIVISION_DESC
							FROM METRO_IT_INSTOCK_DEPT
							WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION, DIVISION_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
							LEFT JOIN
							(SELECT DIVISION, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
							FROM METRO_IT_INSTOCK_DEPT
							WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
									FROM METRO_IT_INSTOCK_DEPT
									WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY') 
										AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
									GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
								AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION)B
							ON A.DIVISION = B.DIVISION ORDER BY 1
						});
				}
				elsif ($tot_wh eq 1){
					$sls3 = $dbh->prepare (qq{
						SELECT A.DIVISION, A.DIVISION_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
							(SELECT DIVISION, DIVISION_DESC
							FROM METRO_IT_INSTOCK_DEPT
							WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION, DIVISION_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
							LEFT JOIN
							(SELECT DIVISION, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
							FROM METRO_IT_INSTOCK_DEPT
							WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
									FROM METRO_IT_INSTOCK_DEPT
									WHERE STORE_FORMAT = 8
									GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
								AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION)B
							ON A.DIVISION = B.DIVISION ORDER BY 1
						});
				}
				elsif ($tot eq 1){
					$sls3 = $dbh->prepare (qq{
						SELECT A.DIVISION, A.DIVISION_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
							(SELECT DIVISION, DIVISION_DESC
							FROM METRO_IT_INSTOCK_DEPT
							WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION, DIVISION_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
							LEFT JOIN
							(SELECT DIVISION, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
							FROM METRO_IT_INSTOCK_DEPT
							WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
									FROM METRO_IT_INSTOCK_DEPT
									WHERE UPPER(REGION) <> 'DUMMY'
										AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
									GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
								AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION)B
							ON A.DIVISION = B.DIVISION ORDER BY 1
						});
				}
				else {
					$sls3 = $dbh->prepare (qq{
						SELECT A.DIVISION, A.DIVISION_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
							(SELECT DIVISION, DIVISION_DESC
							FROM METRO_IT_INSTOCK_DEPT
							WHERE GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION, DIVISION_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
							LEFT JOIN
							(SELECT DIVISION, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
							FROM METRO_IT_INSTOCK_DEPT
							WHERE STORE_CODE IN ('$store') AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
							GROUP BY DIVISION)B
							ON A.DIVISION = B.DIVISION ORDER BY 1
						});
				}
				
				$sls3->execute();				
				$grp_counter = $a;
				while(my $s = $sls3->fetchrow_hashref()){
					$division = $s->{DIVISION};
					$division_desc = $s->{DIVISION_DESC};
					
					if ($tot_vis eq 1){
						$sls4 = $dbh->prepare (qq{	 
							SELECT A.DEPARTMENT_CODE, A.DEPARTMENT_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
								(SELECT DEPARTMENT_CODE, DEPARTMENT_DESC
								FROM METRO_IT_INSTOCK_DEPT
								WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
								LEFT JOIN
								(SELECT DEPARTMENT_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
								FROM METRO_IT_INSTOCK_DEPT
								WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
										FROM METRO_IT_INSTOCK_DEPT
										WHERE UPPER(REGION) LIKE '%VISAYAS%' AND (  (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
										GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
									AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE)B
								ON A.DEPARTMENT_CODE = B.DEPARTMENT_CODE ORDER BY 1
							});
					}
					elsif ($tot_luz eq 1){
						$sls4 = $dbh->prepare (qq{	 
							SELECT A.DEPARTMENT_CODE, A.DEPARTMENT_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
								(SELECT DEPARTMENT_CODE, DEPARTMENT_DESC
								FROM METRO_IT_INSTOCK_DEPT
								WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
								LEFT JOIN
								(SELECT DEPARTMENT_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
								FROM METRO_IT_INSTOCK_DEPT
								WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
										FROM METRO_IT_INSTOCK_DEPT
										WHERE (UPPER(REGION) NOT LIKE '%VISAYAS%' AND UPPER(REGION) <> 'DUMMY') 
											AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
										GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
									AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE)B
								ON A.DEPARTMENT_CODE = B.DEPARTMENT_CODE ORDER BY 1
							});
					}
					elsif ($tot_wh eq 1){
						$sls4 = $dbh->prepare (qq{	 
							SELECT A.DEPARTMENT_CODE, A.DEPARTMENT_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
								(SELECT DEPARTMENT_CODE, DEPARTMENT_DESC
								FROM METRO_IT_INSTOCK_DEPT
								WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
								LEFT JOIN
								(SELECT DEPARTMENT_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
								FROM METRO_IT_INSTOCK_DEPT
								WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
										FROM METRO_IT_INSTOCK_DEPT
										WHERE STORE_FORMAT = 8
										GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
									AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE)B
								ON A.DEPARTMENT_CODE = B.DEPARTMENT_CODE ORDER BY 1
							});
					}
					elsif ($tot eq 1){
						$sls4 = $dbh->prepare (qq{	 
							SELECT A.DEPARTMENT_CODE, A.DEPARTMENT_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
								(SELECT DEPARTMENT_CODE, DEPARTMENT_DESC
								FROM METRO_IT_INSTOCK_DEPT
								WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
								LEFT JOIN
								(SELECT DEPARTMENT_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
								FROM METRO_IT_INSTOCK_DEPT
								WHERE STORE_CODE IN ( SELECT DISTINCT STORE_CODE
										FROM METRO_IT_INSTOCK_DEPT
										WHERE UPPER(REGION) <> 'DUMMY'
											AND ( (STORE_CODE > '1999' AND STORE_CODE < '5000') OR (STORE_CODE > '6000' AND STORE_CODE < '6999') )
										GROUP BY STORE_CODE HAVING SUM(TOT_REPL_ITEMS) <> 0 ) 
									AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE)B
								ON A.DEPARTMENT_CODE = B.DEPARTMENT_CODE ORDER BY 1
							});
					}
					else{
						$sls4 = $dbh->prepare (qq{	 
							SELECT A.DEPARTMENT_CODE, A.DEPARTMENT_DESC, B.TOT_ITEMS, B.REP_ITEMS FROM
								(SELECT DEPARTMENT_CODE, DEPARTMENT_DESC
								FROM METRO_IT_INSTOCK_DEPT
								WHERE DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE, DEPARTMENT_DESC HAVING SUM(TOT_REPL_ITEMS) <> 0)A
								LEFT JOIN
								(SELECT DEPARTMENT_CODE, SUM(TOT_REPL_ITEMS) AS TOT_ITEMS, SUM(REPL_WITH_SOH) AS REP_ITEMS
								FROM METRO_IT_INSTOCK_DEPT
								WHERE STORE_CODE IN ('$store') AND DIVISION = '$division' AND GROUP_CODE = '$group_code' AND MERCH_GROUP_CODE = '$merch_group_code'
								GROUP BY DEPARTMENT_CODE)B
								ON A.DEPARTMENT_CODE = B.DEPARTMENT_CODE ORDER BY 1
							});
					}
					
					$sls4->execute();
					
					while(my $s = $sls4->fetchrow_hashref()){
						
						$worksheet->write($a,5, $s->{DEPARTMENT_CODE},$desc);
						$worksheet->write($a,6, $s->{DEPARTMENT_DESC},$desc);
						
						$worksheet->write($a,$col, $s->{TOT_ITEMS},$border1);
						$worksheet->write($a,$col+1, $s->{REP_ITEMS},$border1);
						
						if ($s->{TOT_ITEMS} <= 0){
							$worksheet->write($a,$col+2, "",$subt);}
						else{
							$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$subt);}
										
						$a++;
						$counter++;
				
					}
					
					$worksheet->write($a,$col, $s->{TOT_ITEMS},$bodyNum);
					$worksheet->write($a,$col+1, $s->{REP_ITEMS},$bodyNum);
					
					if ($s->{TOT_ITEMS} <= 0){
						$worksheet->write($a,$col+2, "",$bodyPct);}
					else{
						$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$bodyPct);}
					
					if($stopper eq 0){
						$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
						$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
					}
					
					$counter = 0;
					$a++;
				}
				
				$worksheet->write($a,$col, $s->{TOT_ITEMS},$bodyNum);
				$worksheet->write($a,$col+1, $s->{REP_ITEMS},$bodyNum);
				
				if ($s->{TOT_ITEMS} <= 0){
					$worksheet->write($a,$col+2, "",$bodyPct);}
				else{
					$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$bodyPct);}
				
				if($stopper eq 0){
					$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
					$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );
				}
				$a++;
			}
			
			if ($merch_group_code eq 'DS' and $stopper eq 0){
				$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
				$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
			}
			
			elsif($merch_group_code eq 'SU' and $stopper eq 0){
				$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
				$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
			}
			
			elsif($merch_group_code eq 'Z_OT' and $stopper eq 0){
				$worksheet->merge_range( $a, 3, $a, 6, 'Total Others', $headN );
				$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'OTHERS', $border2 );
			}
			
			$worksheet->write($a,$col, $s->{TOT_ITEMS},$headNumber);
			$worksheet->write($a,$col+1, $s->{REP_ITEMS},$headNumber);
			
			if ($s->{TOT_ITEMS} <= 0){
				$worksheet->write($a,$col+2, "",$headPct);}
			else{
				$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$headPct);}
			
			$a++;
		}

	$worksheet->write($a,$col, $s->{TOT_ITEMS},$headNumber);
	$worksheet->write($a,$col+1, $s->{REP_ITEMS},$headNumber);
	
	if ($s->{TOT_ITEMS} <= 0){
		$worksheet->write($a,$col+2, "",$headPct);}
	else{
		$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{REP_ITEMS}/$s->{TOT_ITEMS} .'),"",('.$s->{REP_ITEMS}/$s->{TOT_ITEMS}. '))',$headPct);}

}

if($stopper eq 0){	
	$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );
}

if ($tot_vis eq 1 or $tot_luz eq 1 or $tot_wh eq 1 or $tot eq 1){ $worksheet->merge_range( $loc_pt, $col, $loc_pt, $col+2, $loc_code .'-'. $loc_desc, $headN ); }
else{ $worksheet->merge_range( $loc_pt, $col, $loc_pt, $col+2, $loc_code .'-'. $loc_desc, $subhead ); }	

$sls->finish();
$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$counter = 0;

$counter = 0;

}

#mailer
sub mail1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file, $attachment_file_2 ) = @ARGV;

$to = ' arthur.emmanuel@metrogaisano.com,fili.mercado@metrogaisano.com,emily.silverio@metrogaisano.com,ronald.dizon@metrogaisano.com,chit.lazaro@metrogaisano.com,jocelyn.sarmiento@metrogaisano.com,charisse.mancao@metrogaisano.com,cindy.yu@metrogaisano.com,cresilda.dehayco@metrogaisano.com,evan.inocencio@metrogaisano.com,fe.botero@metrogaisano.com,jonrel.nacor@metrogaisano.com,junah.oliveron@metrogaisano.com,lyn.cabatuan@metrogaisano.com,zenda.mangabon@metrogaisano.com,joyce.mirabueno@metrogaisano.com,mariegrace.ong@metrogaisano.com,cherry.gulloy@metrogaisano.com,janice.bedrijo@metrogaisano.com,jerson.roma@metrogaisano.com,bermon.alcantara@metrogaisano.com,nilynn.yosores@metrogaisano.com,anafatima.mancho@metrogaisano.com,emily.silverio@metrogaisano.com,leslie.chipeco@metrogaisano.com,karan.malani@metrogaisano.com, margaret.ang@metrogaisano.com,
gea.badiang@metroretail.com.ph,brein.tajanlangit@metroretail.com.ph,regine.caparida@metroretail.com.ph';

$cc = 'luz.bitang@metrogaisano.com, rex.cabanilla@metrogaisano.com, annalyn.conde@metrogaisano.com,rashel.legaspi@metroretail.com.ph';

$bcc = ' lea.gonzaga@metrogaisano.com;

$from = 'Report Mailer<report.mailer@metrogaisano.com>';		

$subject = 'Replenishment In-stock';

$msgbody_file = 'message.txt';

$attachment_file = "Replenishment In-stock v1.21.xlsx";

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
Content-Type: text/html; charset="iso-8859-1"
Content-Transfer-Encoding: quoted-printable

<html>
Dear Users, <br> <br>
Please refer below for details of the enclosed report: <br> <br>

<table border = 1>
	<tr>
		<th>Measure</th>
		<th>Definition</th>
	</tr>
	<tr>
		<td>SKU Count</td>
		<td>Count of SKUs which satisfies the following conditions:
			<ul><li>Active status (A)</li>
				<li>With replenishment parameter for a particular store or warehouse location</li>
				<li>Orderable flag (Y)</li>
				<li>First Received Date is not NULL</li>
				<li>Last Received Date within the last 3 months (Supermarket) and 6 months (General Merchandise)</li></ul></td>
	</tr>
	<tr>
		<td>In-Stock</td>
		<td>Count of replenishment SKUs with Stock On-hand greater than zero</td>
	</tr>
	<tr>
		<td>% In-Stock</td>
		<td>In-stock / SKU Count</td>
	</tr>
</table>
<br>

If you need assistance, kindly email arcbi.support&#64;metrogaisano.com. <br> <br>

Regards, <br>
ARC BI Support <p>
</html>

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

$to = ' manuel.degamo@metrogaisano.com, ace.olalia@metrogaisano.com, alma.espino@metrogaisano.com, angeli_christi.ladot@metrogaisano.com, angelito.dublin@metrogaisano.com, arlene.yanson@metrogaisano.com, augosto.daria@metrogaisano.com, charm.buenaventura@metrogaisano.com, teena.velasco@metrogaisano.com, cristy.sy@metrogaisano.com, diana.almagro@metrogaisano.com, edgardo.lim@metrogaisano.com, edris.tarrobal@metrogaisano.com, fidela.villamor@metrogaisano.com, genaro.felisilda@metrogaisano.com, genevive.quinones@metrogaisano.com, glenda.navares@metrogaisano.com, joefrey.camu@metrogaisano.com, jonalyn.diaz@metrogaisano.com, opcplanning@metrogaisano.com ';

$cc = 'kent.mamalias@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'Replenishment In-stock';

$msgbody_file = 'message.txt';

$attachment_file = "Replenishment In-stock v1.21.xlsx";

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
Content-Type: text/html; charset="iso-8859-1"
Content-Transfer-Encoding: quoted-printable

<html>
Dear Users, <br> <br>
Please refer below for details of the enclosed report: <br> <br>

<table border = 1>
	<tr>
		<th>Measure</th>
		<th>Definition</th>
	</tr>
	<tr>
		<td>SKU Count</td>
		<td>Count of SKUs which satisfies the following conditions:
			<ul><li>Active status (A)</li>
				<li>With replenishment parameter for a particular store or warehouse location</li>
				<li>Orderable flag (Y)</li>
				<li>First Received Date is not NULL</li>
				<li>Last Received Date within the last 3 months (Supermarket) and 6 months (General Merchandise)</li></ul></td>
	</tr>
	<tr>
		<td>In-Stock</td>
		<td>Count of replenishment SKUs with Stock On-hand greater than zero</td>
	</tr>
	<tr>
		<td>% In-Stock</td>
		<td>In-stock / SKU Count</td>
	</tr>
</table>
<br>

If you need assistance, kindly email arcbi.support&#64;metrogaisano.com. <br> <br>

Regards, <br>
ARC BI Support <p>
</html>

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

$to = ' josemarie.graciadas@metrogaisano.com, jovany.polancos@metrogaisano.com, judy.gilo@metrogaisano.com, julie.montano@metrogaisano.com, kathlene.procianos@metrogaisano.com, limuel.ulanday@metrogaisano.com, cristina.de_asis@metrogaisano.com, mariajoana.cruz@metrogaisano.com, may.sasedor@metrogaisano.com, michelle.calsada@metrogaisano.com, policarpo.mission@metrogaisano.com, rex.refuerzo@metrogaisano.com, ricky.tulda@metrogaisano.com, ronald.dizon@metrogaisano.com, roselle.agbayani@metrogaisano.com, rowena.tangoan@metrogaisano.com, roy.igot@metrogaisano.com, tessie.cabanero@metrogaisano.com, victoria.ferolino@metrogaisano.com, wendel.gallo@metrogaisano.com, juanjose.sibal@metrogaisano.com, julie.montano@metrogaisano.com, kristine.apurado@metrogaisano.com, bebelyn.cabasan@metrogaisano.com, analiza.dano@metrogaisano.com, josemari.abellana@metrogaisano.com ';

$cc = 'kent.mamalias@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'Replenishment In-stock';

$msgbody_file = 'message.txt';

$attachment_file = "Replenishment In-stock v1.21.xlsx";

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
Content-Type: text/html; charset="iso-8859-1"
Content-Transfer-Encoding: quoted-printable

<html>
Dear Users, <br> <br>
Please refer below for details of the enclosed report: <br> <br>

<table border = 1>
	<tr>
		<th>Measure</th>
		<th>Definition</th>
	</tr>
	<tr>
		<td>SKU Count</td>
		<td>Count of SKUs which satisfies the following conditions:
			<ul><li>Active status (A)</li>
				<li>With replenishment parameter for a particular store or warehouse location</li>
				<li>Orderable flag (Y)</li>
				<li>First Received Date is not NULL</li>
				<li>Last Received Date within the last 3 months (Supermarket) and 6 months (General Merchandise)</li></ul></td>
	</tr>
	<tr>
		<td>In-Stock</td>
		<td>Count of replenishment SKUs with Stock On-hand greater than zero</td>
	</tr>
	<tr>
		<td>% In-Stock</td>
		<td>In-stock / SKU Count</td>
	</tr>
</table>
<br>

If you need assistance, kindly email arcbi.support&#64;metrogaisano.com. <br> <br>

Regards, <br>
ARC BI Support <p>
</html>

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








