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

($year,$month,$day) = Today();
$month_to_text = Month_to_Text($month);

my $workbook = Excel::Writer::XLSX->new("Non Replenishment In-stock v1.xlsx");
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

printf "NON REP IN STOCK REPORT \n";

&generate_csv;

&new_sheet($sheet = "Summary-Store");
&call_str_merchandise;

&new_sheet($sheet = "Summary-Division");
&call_summary_division;

&new_sheet($sheet = "Store", $comment = "STORES");		
&call_vis;
&calc_tot_region($total_label = 'TOTAL VISAYAS', $end = 0);
&call_luz;
&calc_tot_region($total_label = 'TOTAL LUZON', $end = 0);
&calc_tot_region($total_label = 'TOTAL', $end = 1);

&new_sheet($sheet = "Warehouse", $comment = "WAREHOUSE");
&call_wh;
&calc_tot_region($total_label = 'TOTAL WAREHOUSE', $end = 0);

$workbook->close();
$dbh_csv->disconnect;


# &mail1;	
# &mail2;	
# &mail3;	
&mail;	

exit;
 
#================================= FUNCTIONS ==================================#

sub call_vis {

$a = 0; $vis = 7; $col = 7; $test = 1; $stopper = 0; 

&heading_orig;

foreach $loc ( '2001', '2002', '2003', '2008', '2010', '2011', '2001W', '4003', '3009', '6001', '6003', '6005','6006','6010', '3001', '3002', '3003', '3004', '3005', '3006', '6004', '6008', '6009' ){
# foreach $loc ( '2001', '2002' ){

	$a += 6, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $len = 0, $loc_pt = $a-2;

	$worksheet->set_column( $col, $col+1, 7 );
	$worksheet->set_column( $col+2, $col+2, 5 );

	foreach my $i ("COUNT", "IN STOCK", "%") {
		$worksheet->write($a-1, $col++, $i, $subhead);
	}

	$col -= 3;

	&query_dept_store($loc);

	$test = 0, $a = 0, $counter = 0, $stopper = 1, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

}

}

sub call_luz {

foreach $loc ( '2004', '2005', '2006', '2007', '2009', '2012', '2013', '2223', '3007', '3012', '4004', '6002', '6012','6013','6011' ){
# foreach $loc ( '2004', '2005' ){

	$a += 6, $counter  =0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $len = 0, $loc_pt = $a-2;

	$worksheet->set_column( $col, $col+1, 7 );
	$worksheet->set_column( $col+2, $col+2, 5 );

	foreach my $i ("COUNT", "IN STOCK", "%") {
		$worksheet->write($a-1, $col++, $i, $subhead);
	}

	$col -= 3;

	&query_dept_store($loc);

	$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

}

}

sub call_wh {

$vis = 7; $col = 7; $test = 1; $stopper = 0; 

&heading_orig;

foreach $loc ( '80001', '80011', '80031', '80041', '80051', '80061', '80141' ){
# foreach $loc ( '80001' ){

	$a += 6, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $len = 0, $loc_pt = $a-2;

	$worksheet->set_column( $col, $col+1, 7 );
	$worksheet->set_column( $col+2, $col+2, 5 );

	foreach my $i ("COUNT", "IN STOCK", "%") {
		$worksheet->write($a-1, $col++, $i, $subhead);
	}

	$col -= 3;

	&query_dept_store($loc);

	$test = 0, $a = 0, $counter = 0, $stopper = 1, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

}

}

sub call_str_merchandise {

$a=10, $d=0, $e=10, $counter=0, $comp_su = 0, $new_su = 0, $comp_nb = 0, $comp_hy = 0, $new_hy = 0, $comp_ds = 0, $new_ds = 0, $type_test = 3;

$worksheet->write($a-9, 3, "Non Replenishment In-Stock", $bold1);
$worksheet->write($a-8, 3, $day . '-' . $month_to_text . '-' .$year, $bold2);

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
#&strNew_Ds;
&strComp_Su;
&strNew_Su;
&strComp_Hy;
&strNew_Hy;
&strComp_Nb;
#&strNew_Nb;

}

sub call_summary_division {

##=============VISAYAS
$a = 0; $vis = 7; $col = 7; $test = 1; $stopper = 0; 

&heading_4;

$loc1 = '2001', $loc2 = '2002', $loc3 = '2003', $loc4 = '2008', $loc5 = '2010', $loc6 = '2011', $loc7 = '2001W', $loc8 = '4003', $loc9 = '3009', $loc10 = '6001', $loc11 = '6003', $loc12 = '6005', $loc13 = '6010', $loc14 = '3001', $loc15 = '3002', $loc16 = '3003', $loc17 = '3004', $loc18 = '3005', $loc19 = '3006', $loc21 = '6004', $loc22 = '6008', $loc23 = '6009', $loc24 = '6006', $loc25 = '0000', $loc26 = '0000', $loc27 = '0000', $loc28 = '0000', $loc29 = '0000', $loc30 = '0000', $loc31 = '0000', $loc32 = '0000', $loc33 = '0000', $loc34 = '0000', $loc35 = '0000', $loc36 = '0000', $loc37 = '0000',$loc38 = '0000',$loc39 = '0000', $region = 'Visayas Stores';
$a += 6, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $stopper = 1, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

##==========LUZON

$loc1 = '2004', $loc2 = '2005', $loc3 = '2006', $loc4 = '2007', $loc5 = '2009', $loc6 = '2012', $loc7 = '2013', $loc8 = '2223', $loc9 = '3007', $loc10 = '4004', $loc11 = '6002', $loc12 = '6012', $loc13 = '3012', $loc14 = '6013', $loc15 = '6011', $loc16 = '0000', $loc17 = '0000', $loc18 = '0000', $loc19 = '0000', $loc21 = '0000', $loc22 = '0000', $loc23 = '0000', $loc24 = '0000', $loc25 = '0000', $loc26 = '0000', $loc27 = '0000', $loc28 = '0000', $loc29 = '0000', $loc30 = '0000', $loc31 = '0000', $loc32 = '0000', $loc33 = '0000', $loc34 = '0000', $loc35 = '0000', $loc36 = '0000',$loc37 = '0000',$loc38 = '0000',$loc39 = '0000', $region = 'Luzon Stores';
$a += 6, $counter  =0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

##==========VISAYAS+LUZON

$loc1 = '2001', $loc2 = '2002', $loc3 = '2003', $loc4 = '2008', $loc5 = '2010', $loc6 = '2011', $loc7 = '2001W', $loc8 = '4003', $loc9 = '3009', $loc10 = '6001', $loc11 = '6003', $loc12 = '6005', $loc13 = '6010', $loc14 = '3001', $loc15 = '3002', $loc16 = '3003', $loc17 = '3004', $loc18 = '3005', $loc19 = '3006', $loc21 = '6004', $loc22 = '6008', $loc23 = '2004', $loc24 = '2005', $loc25 = '2006', $loc26 = '2007', $loc27 = '2009', $loc28 = '2012', $loc29 = '2013', $loc30 = '2223', $loc31 = '3007', $loc32 = '4004', $loc33 = '6002', $loc34 = '6012', $loc35 = '3012', $loc36 = '6009', $loc37 = '6013',$loc38 = '6006',$loc39 = '0000',$region = 'All Stores';
$a += 6, $counter  =0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

##==========WAREHOUSE

$loc1 = '80001', $loc2 = '80011', $loc3 = '80031', $loc4 = '80041', $loc5 = '80051', $loc6 = '80061', $loc7 = '80141', $loc8 = '0000', $loc9 = '0000', $loc10 = '0000', $loc11 = '0000', $loc12 = '0000', $loc13 = '0000', $loc14 = '0000', $loc15 = '0000', $loc16 = '0000', $loc17 = '0000', $loc18 = '0000', $loc19 = '0000', $loc21 = '0000', $loc22 = '0000', $loc23 = '0000', $loc24 = '0000', $loc25 = '0000', $loc26 = '0000', $loc27 = '0000', $loc28 = '0000', $loc29 = '0000', $loc30 = '0000', $loc31 = '0000', $loc32 = '0000', $loc33 = '0000', $loc34 = '0000', $loc35 = '0000', $loc36 = '0000', $loc37 = '0000',$region = 'All Warehouses';
$a += 6, $counter = 0, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $stopper = 1, $total_tot_items = 0, $total_rep_items = 0, $grp_tot_items = 0, $grp_rep_items = 0, $col += 3;

}


sub strComp_Su {

$div_name = "Comp";  $div_name3 = "Supermarket";
$mrch1 = 'SU'; $mrch2 = 'SU'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8040'; $dept_grp2 = '0000'; $dept_grp3 = '0000'; $dept_grp4 = '0000'; $dept_grp5 = '0000'; $dept_grp6 = '0000'; $dept_grp7 = '0000';
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU2001'; $store2 = 'SU2002'; $store3 = 'SU2003'; $store4 = 'SU2004'; $store5 = 'SU2006'; $store6 = 'SU2007'; $store7 = 'SU2009'; $store8 = 'SU2012'; $store9 = 'SU2001W'; $store10 = 'SU2013'; $store11 = 'SU4004'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000'; $store41 = '0000'; $store42 = '0000'; $store43 = '0000'; $store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2012'; $stor9 = '2001W'; $stor10 = '2013'; $stor11 = '4004'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000';   $stor15 = '0000'; $stor16 = '0000'; 

	if($type_test eq 3){	&query_summary_merchandise;	} 
	
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
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = '0000'; $store2 = '0000'; $store3 = 'SU3009'; $store4 = 'SU3010'; $store5 = 'SU3011'; $store6 = 'SU3012'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000'; $store41 = '0000'; $store42 = '0000'; $store43 = '0000'; $store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '0000'; $stor2 = '0000'; $stor3 = '3009'; $stor4 = '3010'; $stor5 = '3011'; $stor6 = '3012'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000';    $stor15 = '0000'; $stor16 = '0000'; 

	if($type_test eq 3){	&query_summary_merchandise;	} 
	
	elsif($type_test eq 4) {	
		
		&query_by_store_merchandise;
		&calc8;

		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_su=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 11, 13, 14 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_su,$col).','.xl_rowcol_to_cell($new_su,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 11 or $col eq 14){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct );
						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
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
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU3001'; $store2 = 'SU3007'; $store3 = 'SU4003'; $store4 = 'DS3001'; $store5 = 'DS3007'; $store6 = 'DS4003'; $store7 = 'OT3001'; $store8 = 'OT3007'; $store9 = 'OT4003'; $store10 = 'SU3002'; $store11 = 'DS3002'; $store12 = 'OT3002'; $store13 = 'SU3003'; $store14 = 'DS3003'; $store15 = 'OT3003'; $store16 = 'SU3004'; $store17 = 'DS3004'; $store18 = 'OT3004'; $store19 = 'SU3005'; $store20 = 'DS3005'; $store21 = 'OT3005'; $store22 = 'SU3006'; $store23 = 'DS3006'; $store24 = 'OT3006'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000'; $store41 = '0000'; $store42 = '0000'; $store43 = '0000';$store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '30001'; $stor2 = '3007'; $stor3 = '4003'; $stor4 = '3002'; $stor5 = '3003'; $stor6 = '3004'; $stor7 = '3005'; $stor8 = '3006'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000';    $stor15 = '0000'; $stor16 = '0000'; 

	if($type_test eq 3){	
	
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

			foreach my $col( 7, 8, 10, 11, 13, 14 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 11 or $col eq 14){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct );
						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
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
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = '0000'; $store2 = '0000'; $store3 = '0000'; $store4 = '0000'; $store5 = '0000'; $store6 = '0000'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000'; $store41 = '0000'; $store42 = '0000'; $store43 = '0000';$store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '0000'; $stor2 = '0000'; $stor3 = '0000'; $stor4 = '0000'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000';     $stor15 = '0000'; $stor16 = '0000'; 

	if($type_test eq 3){	
	
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

		$new_nb=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 11, 13, 14 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_nb,$col).','.xl_rowcol_to_cell($new_nb,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 11 or $col eq 14){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct );
						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
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
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU2005'; $store2 = 'SU2008'; $store3 = 'SU2010'; $store4 = 'SU2011'; $store5 = 'DS2005'; $store6 = 'DS2008'; $store7 = 'DS2010'; $store8 = 'DS2011'; $store9 = 'DS2005'; $store10 = 'OT2005'; $store11 = 'OT2008'; $store12 = 'OT2010'; $store13 = 'OT2011'; $store14 = 'SU6001'; $store15 = 'DS6001'; $store16 = 'OT6001'; $store17 = 'SU6002'; $store18 = 'DS6002'; $store19 = 'OT6002'; $store20 = 'SU6003'; $store21 = 'DS6003'; $store22 = 'OT6003'; $store23 = 'SU6005'; $store24 = 'DS6005'; $store25 = 'OT6005'; $store26 = 'SU6010';  $store27 = 'DS6010'; $store28 = 'OT6010'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000'; $store41 = '0000'; $store42 = '0000'; $store43 = '0000';$store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '2005'; $stor2 = '2008'; $stor3 = '2010'; $stor4 = '2011'; $stor5 = '6001'; $stor6 = '6002'; $stor7 = '6003'; $stor8 = '6005'; $stor9 = '6010'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000';     $stor15 = '0000';  $stor16 = '0000'; 

	if($type_test eq 3){	&query_summary_merchandise;	} 
	
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
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = '0000'; $store2 = '0000'; $store3 = '0000'; $store4 = 'SU6004'; $store5 = '0000'; $store6 = 'SU6012'; $store7 = 'SU6009'; $store8 = 'SU6013'; $store9 = 'SU6011'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = 'DS6004'; $store14 = '0000'; $store15 = 'DS6012'; $store16 = 'DS6009'; $store17 = 'DS6013'; $store18 = 'DS6011'; $store19 = '0000'; $store20 = '0000'; $store21 = 'OT6004'; $store22 = '0000'; $store23 = 'OT6012'; $store24 = 'OT6009'; $store25 = 'OT6013'; $store26 = 'OT6011';  $store27 = 'OT6000'; $store28 = 'DS6006'; $store29 = 'SU6006'; $store30 = 'OT6006'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';$store41 = '0000'; $store42 = '0000'; $store43 = '0000';$store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '0000'; $stor2 = '0000'; $stor3 = '0000'; $stor4 = '6004'; $stor5 = '0000'; $stor6 = '6012'; $stor7 = '6009'; $stor8 = '6013'; $stor9 = '6011'; $stor10 = '6006'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';   $stor14 = '0000';     $stor15 = '0000'; $stor16 = '0000'; 

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

			foreach my $col( 7, 8, 10, 11, 13, 14 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_hy,$col).','.xl_rowcol_to_cell($new_hy,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 11 or $col eq 14){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct );
						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
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
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'DS2001'; $store2 = 'DS2002'; $store3 = 'DS2003'; $store4 = 'DS2004'; $store5 = 'DS2006'; $store6 = 'DS2007'; $store7 = 'DS2009'; $store8 = 'DS2223'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';$store41 = '0000'; $store42 = '0000'; $store43 = '0000';$store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2223'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000';    $stor15 = '0000'; $stor16 = '0000'; 

	if($type_test eq 3){	&query_summary_merchandise;	} 
	
	elsif($type_test eq 4) {	
		
		&query_by_store_merchandise;	
		&calc8;
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $a-$counter, 3, $a+1, 3, $div_name3, $border2 );
		
		$comp_ds=$a; $tst = $a-$counter; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER
		
			foreach my $col( 7, 8, 10, 11, 13, 14 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_ds,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 11 or $col eq 14){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct );
						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
						}
					}
			}

		$worksheet->merge_range( $a, 4, $a, 6, "Total ". $div_name3, $bodyN );
		$a+=1; $counter = 0; $d=$a;
	}

}

sub strNew_Ds {

$div_name = "New"; $div_name2 = "Department Store"; $div_name3 = "Department Store";
$mrch1 = 'DS'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '9000'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8050'; $dept_grp5 = '8060'; $dept_grp6 = '8070'; $dept_grp7 = '0000';
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = '0000'; $store2 = '0000'; $store3 = '0000'; $store4 = '0000'; $store5 = '0000'; $store6 = '0000'; $store7 = '0000'; $store8 = '0000'; $store9 = '0000'; $store10 = '0000'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';$store41 = '0000'; $store42 = '0000'; $store43 = '0000';$store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '0000'; $stor2 = '0000'; $stor3 = '0000'; $stor4 = '0000'; $stor5 = '0000'; $stor6 = '0000'; $stor7 = '0000'; $stor8 = '0000'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000';    $stor15 = '0000'; $stor16 = '0000'; 

	if($type_test eq 3){	&query_summary_merchandise;	}	
	
	elsif($type_test eq 4) {	
		
		&query_by_store_merchandise;	
		&calc8;
		
		$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
		$worksheet->merge_range( $a-$counter, 4, $a, 4, $div_name, $border2 );
		$worksheet->merge_range( $tst, 3, $a+1, 3, $div_name2, $border2 );

		$new_ds=$a; $a+=1; $b+=1; $d=$a; $counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER

			foreach my $col( 7, 8, 10, 11, 13, 14 ){
				my $sumTY = '=SUM('.xl_rowcol_to_cell($comp_ds,$col).','.xl_rowcol_to_cell($new_ds,$col).')';
					$worksheet->write( $a, $col, $sumTY, $bodyNum );
					
					if ($col eq 8 or $col eq 11 or $col eq 14){
						if( xl_rowcol_to_cell( $a, $col ) le 0){
							$worksheet->write( $a, $col+1, '', $bodyPct );
						}
						else{
							my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col ). '/' . xl_rowcol_to_cell( $a, $col-1 ) . ')';
							$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
						}
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
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;

$store1 = 'DS2001'; $store2 = 'DS2002'; $store3 = 'DS2003'; $store4 = 'DS2004'; $store5 = 'DS2006'; $store6 = 'DS2007'; $store7 = 'DS2009'; $store8 = 'DS2223'; $store9 = '0000'; $store10 = 'OOOO'; $store11 = '0000'; $store12 = '0000'; $store13 = '0000'; $store14 = '0000'; $store15 = '0000'; $store16 = '0000'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';$store41 = '0000'; $store42 = '0000'; $store43 = '0000';$store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2223'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000';    $stor15 = '0000'; $stor16 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 
	elsif($type_test eq 3){	&query_summary_merchandise;	} 

}

sub str_Su {

$div_name = "Comp"; $div_name3 = 'Supermarket';
$mrch1 = 'SU'; $mrch2 = 'SU'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '0000';
$dept_grp1 = '8040'; $dept_grp2 = '0000'; $dept_grp3 = '0000'; $dept_grp4 = '0000'; $dept_grp5 = '0000'; $dept_grp6 = '0000'; $dept_grp7 = '0000';
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;

$store1 = 'SU2001'; $store2 = 'SU2002'; $store3 = 'SU2003'; $store4 = 'SU2004'; $store5 = 'SU2001W'; $store6 = 'SU2006'; $store7 = 'SU2007'; $store8 = '0000'; $store9 = 'SU2009'; $store10 = 'SU2013'; $store11 = 'SU4004'; $store12 = 'SU2012'; $store13 = 'SU3009'; $store14 = 'SU3010'; $store15 = 'SU3011'; $store16 = 'SU3012'; $store17 = '0000'; $store18 = '0000'; $store19 = '0000'; $store20 = '0000'; $store21 = '0000'; $store22 = '0000'; $store23 = '0000'; $store24 = '0000'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';$store41 = '0000'; $store42 = '0000'; $store43 = '0000';$store44 = '0000'; $store45 = '0000'; $store46 = '0000';

$stor1 = '2001'; $stor2 = '2002'; $stor3 = '2003'; $stor4 = '2004'; $stor5 = '2006'; $stor6 = '2007'; $stor7 = '2009'; $stor8 = '2013'; $stor9 = '2012'; $stor10 = '4004'; $stor11 = '3009'; $stor12 = '3010';  $stor13 = '3011';    $stor14 = '2001W';    $stor15 = '3012'; $stor16 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 
	elsif($type_test eq 3){	&query_summary_merchandise;	} 
		
}

sub str_Hy {

$div_name = "New"; $div_name3 = 'Hypermarket';
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000'; $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU6001'; $store2 = 'SU6002'; $store3 = 'SU6003'; $store4 = 'SU6004'; $store5 = 'SU6005'; $store6 = 'SU6012'; $store7 = 'SU6009'; $store8 = 'SU6010'; $store9 = 'SU6011'; $store10 = 'DS6001'; $store11 = 'DS6002'; $store12 = 'DS6003'; $store13 = 'DS6004'; $store14 = 'DS6005'; $store15 = 'DS6012'; $store16 = 'DS6009'; $store17 = 'DS6010'; $store18 = 'DS6011'; $store19 = 'OT6002'; $store20 = 'OT6003'; $store21 = 'OT6004'; $store22 = 'OT6005'; $store23 = 'OT6012'; $store24 = 'OT6009'; $store25 = 'OT6010'; $store26 = 'OT6011';  $store27 = 'OT6000'; $store28 = 'SU2005'; $store29 = 'SU2008'; $store30 = 'SU2010'; $store31 = 'SU2011'; $store32 = 'DS2005'; $store33 = 'DS2008'; $store34 = 'DS2010'; $store35 = 'DS2011'; $store36 = 'DS2005'; $store37 = 'OT2005'; $store38 = 'OT2008'; $store39 = 'OT2010'; $store40 = 'OT2011';$store41 = 'SU6013'; $store42 = 'DS6013'; $store43 = 'OT6013';$store44 = 'SU6006'; $store45 = 'DS6006'; $store46 = 'OT6006';

$stor1 = '6001'; $stor2 = '6002'; $stor3 = '6003'; $stor4 = '6004'; $stor5 = '6005'; $stor6 = '6012'; $stor7 = '6009'; $stor8 = '6010'; $stor9 = '6011'; $stor10 = '6013'; $stor11 = '6006'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000';    $stor15 = '0000'; $stor16 = '0000'; 

	if($type_test eq 1){	&query_summary;	} 
	elsif($type_test eq 3){	&query_summary_merchandise;	} 

}

sub str_Nb {

$div_name = "All"; $div_name3 = 'Neighborhood Store';
$mrch1 = 'SU'; $mrch2 = 'DS'; $mrch3 = 'OT'; 
$division_grp1 = '8500'; $division_grp2 = '8000';  $division_grp3 = '9000';
$dept_grp1 = '8010'; $dept_grp2 = '8020'; $dept_grp3 = '8030'; $dept_grp4 = '8040'; $dept_grp5 = '8050'; $dept_grp6 = '8060'; $dept_grp7 = '8070';
$s1f2_tot_items = 0, $s1f2_rep_items = 0, $s1f2_counter = 0, $s1f2_row = 0;

#$fmt1 = 1; $fmt2 = 2; $fmt3 = 3; $fmt4 = 4; $fmt5 = 5;
$store1 = 'SU3001'; $store2 = 'SU3007'; $store3 = 'SU4003'; $store4 = 'DS3001'; $store5 = 'DS3007'; $store6 = 'DS4003'; $store7 = 'OT3001'; $store8 = 'OT3007'; $store9 = 'OT4003'; $store10 = 'SU3002'; $store11 = 'DS3002'; $store12 = 'OT3002'; $store13 = 'SU3003'; $store14 = 'DS3003'; $store15 = 'OT3003'; $store16 = 'SU3004'; $store17 = 'DS3004'; $store18 = 'OT3004'; $store19 = 'SU3005'; $store20 = 'DS3005'; $store21 = 'OT3005'; $store22 = 'SU3006'; $store23 = 'DS3006'; $store24 = 'OT3006'; $store25 = '0000'; $store26 = '0000';  $store27 = '0000'; $store28 = '0000'; $store29 = '0000'; $store30 = '0000'; $store31 = '0000'; $store32 = '0000'; $store33 = '0000'; $store34 = '0000'; $store35 = '0000'; $store36 = '0000'; $store37 = '0000'; $store38 = '0000'; $store39 = '0000'; $store40 = '0000';

$stor1 = '30001'; $stor2 = '3007'; $stor3 = '4003'; $stor4 = '3002'; $stor5 = '3003'; $stor6 = '3004'; $stor7 = '3005'; $stor8 = '3006'; $stor9 = '0000'; $stor10 = '0000'; $stor11 = '0000'; $stor12 = '0000';  $stor13 = '0000';    $stor14 = '0000';    $stor15 = '0000'; $stor16 = '0000'; 

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

sub heading_orig {

$worksheet->write(0, 2, "Non Replenishment In-Stock - " . $comment , $bold);
$worksheet->write(1, 2, $day . '-' . $month_to_text . '-' .$year, $bold2);
$worksheet->merge_range( 4, 2, 5, 2, 'Type', $subhead );
$worksheet->merge_range( 4, 3, 5, 3, 'Type', $subhead );
$worksheet->merge_range( 4, 4, 5, 4, 'Type', $subhead );
$worksheet->merge_range( 4, 5, 5, 5, 'Code', $subhead );
$worksheet->merge_range( 4, 6, 5, 6, 'Desc', $subhead );

}

sub heading_3 {

$worksheet->write($a-3, 3, "in 000's", $script);

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

$worksheet->write(0, 2, "Non Replenishment In-Stock " . $comment , $bold);
$worksheet->write(1, 2, $day . '-' . $month_to_text . '-' .$year, $bold2);
$worksheet->merge_range( 4, 2, 5, 2, 'Type', $subhead );
$worksheet->merge_range( 4, 3, 5, 3, 'Type', $subhead );
$worksheet->merge_range( 4, 4, 5, 6, 'Desc', $subhead );

}


sub query_summary_merchandise {

$table = 'instock_nonrep_v1.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;

$sls = $dbh_csv->prepare (qq{SELECT SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (group_no = '$dept_grp1' or group_no = '$dept_grp2' 
																										or group_no = '$dept_grp3' or group_no = '$dept_grp4' 
																										or group_no = '$dept_grp5' or group_no = '$dept_grp6' 
																										or group_no = '$dept_grp7')) or
																		division = '$division_grp3'))
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14' or store_code = '$stor15'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40' or store = '$store41' or store = '$store42' or store = '$store43' 
									or store = '$store44' or store = '$store45' or store = '$store46')
								});
$sls->execute();

if ($mrch1 eq 'DS' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){

	while(my $s = $sls->fetchrow_hashref()){
		$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
		$worksheet->write($a,7, $s->{tot_items},$border1);
		$worksheet->write($a,8, $s->{rep_items},$border1);
		
		if ($s->{tot_items} <= 0){
			$worksheet->write($a,9, "",$subt);}
		else{
			$worksheet->write($a,9, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
			
		$worksheet->write($a,10, "",$border1);
		$worksheet->write($a,11, "",$border1);
		$worksheet->write($a,12, "",$border1);
		
		$worksheet->write($a,13, $s->{tot_items},$border1);
		$worksheet->write($a,14, $s->{rep_items},$border1);
		
		if ($s->{tot_items} <= 0){
			$worksheet->write($a,15, "",$subt);}
		else{
			$worksheet->write($a,15, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
		
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
		
		$worksheet->write($a,10, $s->{tot_items},$border1);
		$worksheet->write($a,11, $s->{rep_items},$border1);
		
		if ($s->{tot_items} <= 0){
			$worksheet->write($a,12, "",$subt);}
		else{
			$worksheet->write($a,12, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
		
		$worksheet->write($a,13, $s->{tot_items},$border1);
		$worksheet->write($a,14, $s->{rep_items},$border1);
		
		if ($s->{tot_items} <= 0){
			$worksheet->write($a,15, "",$subt);}
		else{
			$worksheet->write($a,15, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
		
		$a++;
		$counter++;
	}
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){
	$worksheet->merge_range( $a, 4, $a, 6, $div_name3, $desc );
	
	$sls_2 = $dbh_csv->prepare (qq{SELECT merch_group_code_rev, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (group_no = '$dept_grp1' or group_no = '$dept_grp2' 
																										or group_no = '$dept_grp3' or group_no = '$dept_grp4' 
																										or group_no = '$dept_grp5' or group_no = '$dept_grp6' 
																										or group_no = '$dept_grp7')) or
																		division = '$division_grp3'))
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14' or store_code = '$stor15'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40'  or store = '$store41' or store = '$store42' or store = '$store43'
									or store = '$store44' or store = '$store45' or store = '$store46')
								GROUP BY merch_group_code_rev
								ORDER BY merch_group_code_rev
								});
	$sls_2->execute();

	while(my $s = $sls_2->fetchrow_hashref()){
	
		if($s->{merch_group_code_rev} eq 'DS'){
			$worksheet->write($a,7, $s->{tot_items},$border1);
			$worksheet->write($a,8, $s->{rep_items},$border1);
			
			if ($s->{tot_items} <= 0){
				$worksheet->write($a,9, "",$subt);}
			else{
				$worksheet->write($a,9, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
		}
		
		else{		
			$worksheet->write($a,10, $s->{tot_items},$border1);
			$worksheet->write($a,11, $s->{rep_items},$border1);
			
			if ($s->{tot_items} <= 0){
				$worksheet->write($a,12, "",$subt);}
			else{
				$worksheet->write($a,12, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
		}
		
		$tot_items += $s->{tot_items};
		$rep_items += $s->{rep_items};		
	}
	
	$worksheet->write($a,13, $tot_items,$border1);
	$worksheet->write($a,14, $rep_items,$border1);
			
	if ($tot_items <= 0){
		$worksheet->write($a,15, "",$subt);}
	else{
		$worksheet->write($a,15, '=IF(ISERROR('. $rep_items/$tot_items .'),"",('.$rep_items/$tot_items. '))',$subt);}
				
	$sls_2->finish();
	
	$tot_items = 0;
	$rep_items = 0;
	$a++;
	$counter++;
}

$sls->finish();

}

sub query_by_store_merchandise {

$table = 'instock_nonrep_v1.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;

$blank = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, hidden => 1 );
$worksheet->conditional_formatting( 'H44:M60', { type     => 'cell',  criteria => '=', value    => 0, format   => $blank });	
$worksheet->conditional_formatting( 'F9:Y2000',  { type => 'cell', criteria => '<', value => 0, format => $down } );		

$sls = $dbh_csv->prepare (qq{SELECT store_code, store_name, merch_group_code_rev, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (group_no = '$dept_grp1' or group_no = '$dept_grp2' 
																										or group_no = '$dept_grp3' or group_no = '$dept_grp4' 
																										or group_no = '$dept_grp5' or group_no = '$dept_grp6' 
																										or group_no = '$dept_grp7')) or
																		division = '$division_grp3')) 
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14' or store_code = '$stor15'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40'  or store = '$store41' or store = '$store42' or store = '$store43'
									or store = '$store44' or store = '$store45' or store = '$store46')
								 GROUP BY store_code, store_name, merch_group_code_rev
								 ORDER BY store_code, merch_group_code_rev
								});
$sls->execute();

if ($mrch1 eq 'DS' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){	
	while(my $s = $sls->fetchrow_hashref()){
			$worksheet->write($a,5, $s->{store_code},$desc);
			$worksheet->write($a,6, $s->{store_name},$desc);
			
			if($s->{merch_group_code_rev} eq 'DS'){
				$worksheet->write($a,7, $s->{tot_items},$border1);
				$worksheet->write($a,8, $s->{rep_items},$border1);
				
				if ($s->{tot_items} <= 0){
					$worksheet->write($a,9, "",$subt);}
				else{
					$worksheet->write($a,9, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
					
				$worksheet->write($a,10, "",$border1);
				$worksheet->write($a,11, "",$border1);
				$worksheet->write($a,12, "",$border1);
			}
			
			else{	
				$a -= 1;
				$counter -= 1;
							
				$worksheet->write($a,10, $s->{tot_items},$border1);
				$worksheet->write($a,11, $s->{rep_items},$border1);
				
				if ($s->{tot_items} <= 0){
					$worksheet->write($a,12, "",$subt);}
				else{
					$worksheet->write($a,12, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
			}
			
			$worksheet->write($a,13, '=SUM('. xl_rowcol_to_cell( $a, 7 ). ',' . xl_rowcol_to_cell( $a, 10 ) . ')',$border1);
			$worksheet->write($a,14, '=SUM('. xl_rowcol_to_cell( $a, 8 ). ',' . xl_rowcol_to_cell( $a, 11 ) . ')',$border1);
			$worksheet->write($a,15, '=IF(ISERROR('. xl_rowcol_to_cell( $a, 14 ) . '/' . xl_rowcol_to_cell( $a, 13 ) .'),"",('. xl_rowcol_to_cell( $a, 14 ) . '/' . xl_rowcol_to_cell( $a, 13 ) . '))',$subt);
			
			$a++;
			$counter++;			
	}
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'SU' and $mrch3 eq 'OT' ){	
	
	while(my $s = $sls->fetchrow_hashref()){
		
		if($s1f2_counter ne 2){
			$worksheet->write($a,5, $s->{store_code},$desc);
			$worksheet->write($a,6, $s->{store_name},$desc);
			
			if($s->{merch_group_code_rev} eq 'DS'){ 		
				$worksheet->write($a,7, $s->{tot_items},$border1);
				$worksheet->write($a,8, $s->{rep_items},$border1);
				
				if ($s->{tot_items} <= 0){
					$worksheet->write($a,9, "",$subt);}
				else{
					$worksheet->write($a,9, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}

				$a -= 1;
				$counter -= 1;
			}
			
			else{
				$worksheet->write($a,10, $s->{tot_items},$border1);
				$worksheet->write($a,11, $s->{rep_items},$border1);
				
				if ($s->{tot_items} <= 0){
					$worksheet->write($a,12, "",$subt);}
				else{
					$worksheet->write($a,12, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}

					if ($s->{merch_group_code_rev} eq 'SU' and ($s->{store_code} eq '2001' or $s->{store_code} eq '2001W')) {
						$s1f2_tot_items += $s->{tot_items};
						$s1f2_rep_items += $s->{rep_items};
						$s1f2_counter ++; # once value = 2, we'll have a summation of s1 and f2						
				}
			
			$worksheet->write($a,13, '=SUM('. xl_rowcol_to_cell( $a, 7 ). ',' . xl_rowcol_to_cell( $a, 10 ) . ')',$border1);
			$worksheet->write($a,14, '=SUM('. xl_rowcol_to_cell( $a, 8 ). ',' . xl_rowcol_to_cell( $a, 11 ) . ')',$border1);
			$worksheet->write($a,15, '=IF(ISERROR('. xl_rowcol_to_cell( $a, 14 ) . '/' . xl_rowcol_to_cell( $a, 13 ) .'),"",('. xl_rowcol_to_cell( $a, 14 ) . '/' . xl_rowcol_to_cell( $a, 13 ) . '))',$subt);
			}
			
			$a++;
			$counter++;			
		}
		
		if($s1f2_counter eq 2){
			$worksheet->write($a,5, "",$desc);
			$worksheet->write($a,6, "METRO COLON + F2",$desc);
			
			$worksheet->write($a,10, $s1f2_tot_items,$border1);
			$worksheet->write($a,11, $s1f2_rep_items,$border1);
			
			if ($s1f2_tot_items <= 0){
				$worksheet->write($a,12, "",$subt);}
			else{
				$worksheet->write($a,12, '=IF(ISERROR('. $s1f2_rep_items/$s1f2_tot_items .'),"",('.$s1f2_rep_items/$s1f2_tot_items. '))',$subt);}
			
			$worksheet->set_row( $a, undef, undef, 1, undef, undef ); #we hide this row
			
			$s1f2_row = $a;
			$s1f2_counter = 0;
			
			$a++;
			$counter++;
		}
		
	}
	
}

elsif ($mrch1 eq 'SU' and $mrch2 eq 'DS' and $mrch3 eq 'OT' ){
	
	$sls_2 = $dbh_csv->prepare (qq{SELECT store_code, store_name, merch_group_code_rev, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
								 FROM $table 
								 WHERE ((merch_group_code = '$mrch3' and (division = '$division_grp1' or 
																		(division = '$division_grp2' and (group_no = '$dept_grp1' or group_no = '$dept_grp2' 
																										or group_no = '$dept_grp3' or group_no = '$dept_grp4' 
																										or group_no = '$dept_grp5' or group_no = '$dept_grp6' 
																										or group_no = '$dept_grp7')) or
																		division = '$division_grp3')) 
										and (store_code = '$stor1' or store_code = '$stor2' or store_code = '$stor3' or store_code = '$stor4' or store_code = '$stor5' 
										or store_code = '$stor6' or store_code = '$stor7' or store_code = '$stor8' or store_code = '$stor9' or store_code = '$stor10'  
										or store_code = '$stor11' or store_code = '$stor12' or store_code = '$stor13' or store_code = '$stor14' or store_code = '$stor15'))
									OR (store = '$store1' or store = '$store2' or store = '$store3' or store = '$store4' or store = '$store5' or store = '$store6' 
									or store = '$store7' or store = '$store8' or store = '$store9' or store = '$store10' or store = '$store11' or store = '$store12' 
									or store = '$store13' or store = '$store14' or store = '$store15' or store = '$store16' or store = '$store17' or store = '$store18' 
									or store = '$store19' or store = '$store20' or store = '$store21' or store = '$store22' or store = '$store23' or store = '$store24' 
									or store = '$store25' or store = '$store26' or store = '$store27' or store = '$store28' or store = '$store29' or store = '$store30' 
									or store = '$store31' or store = '$store32' or store = '$store33' or store = '$store34' or store = '$store35' or store = '$store36' 
									or store = '$store37' or store = '$store38' or store = '$store39' or store = '$store40'  or store = '$store41' or store = '$store42' or store = '$store43'
									or store = '$store44' or store = '$store45' or store = '$store46')
								 GROUP BY store_code, store_name , merch_group_code_rev
								 ORDER BY store_code, merch_group_code_rev
								});
	$sls_2->execute();	

	while(my $s = $sls_2->fetchrow_hashref()){
	
	$worksheet->write($a,5, $s->{store_code},$desc);
	$worksheet->write($a,6, $s->{store_name},$desc);
	
		if($s->{merch_group_code_rev} eq 'DS'){
			$worksheet->write($a,7, $s->{tot_items},$border1);
			$worksheet->write($a,8, $s->{rep_items},$border1);
			
			if ($s->{tot_items} <= 0){
				$worksheet->write($a,9, "",$subt);}
			else{
				$worksheet->write($a,9, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
		}
		
		else{	
			$worksheet->write($a,10, $s->{tot_items},$border1);
			$worksheet->write($a,11, $s->{rep_items},$border1);
			
			if ($s->{tot_items} <= 0){
				$worksheet->write($a,12, "",$subt);}
			else{
				$worksheet->write($a,12, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);}
		
			$a++;
			$counter++;
		}
		
		$worksheet->write($a,13, '=SUM('. xl_rowcol_to_cell( $a, 7 ). ',' . xl_rowcol_to_cell( $a, 10 ) . ')',$border1);
		$worksheet->write($a,14, '=SUM('. xl_rowcol_to_cell( $a, 8 ). ',' . xl_rowcol_to_cell( $a, 11 ) . ')',$border1);
		$worksheet->write($a,15, '=IF(ISERROR('. xl_rowcol_to_cell( $a, 14 ) . '/' . xl_rowcol_to_cell( $a, 13 ) .'),"",('. xl_rowcol_to_cell( $a, 14 ) . '/' . xl_rowcol_to_cell( $a, 13 ) . '))',$subt);
	
	}
	
	$sls_2->finish();
	
}

$sls->finish();

}

sub query_summary_division {

$table = 'instock_nonrep_v1.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls1 = $dbh_csv->prepare (qq{SELECT merch_group_code_rev, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
								 FROM $table 
								 WHERE (store_code = '$loc1' or store_code = '$loc2' or store_code = '$loc3' or store_code = '$loc4' or store_code = '$loc5' or store_code = '$loc6' 
									or store_code = '$loc7' or store_code = '$loc8' or store_code = '$loc9' or store_code = '$loc10' or store_code = '$loc11' 
									or store_code = '$loc12' or store_code = '$loc13' or store_code = '$loc14' or store_code = '$loc15' or store_code = '$loc16' 
									or store_code = '$loc17' or store_code = '$loc18' or store_code = '$loc19' or store_code = '$loc20' or store_code = '$loc21' 
									or store_code = '$loc22' or store_code = '$loc23' or store_code = '$loc24' or store_code = '$loc25' or store_code = '$loc26' 
									or store_code = '$loc27' or store_code = '$loc28' or store_code = '$loc29' or store_code = '$loc30' or store_code = '$loc31' 
									or store_code = '$loc32' or store_code = '$loc33' or store_code = '$loc34' or store_code = '$loc35' or store_code = '$loc36'
									or store_code = '$loc37')
									and (division <> '1500' and division <> '8000' and division <> '8500' and division <> '9000')
								 GROUP BY merch_group_code_rev 
								 ORDER BY merch_group_code_rev
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{merch_group_code_rev};
	
	$sls2 = $dbh_csv->prepare (qq{SELECT attrib1, attrib2, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
								 FROM $table 
								 WHERE merch_group_code_rev = '$merch_group_code' and 
									(store_code = '$loc1' or store_code = '$loc2' or store_code = '$loc3' or store_code = '$loc4' or store_code = '$loc5' or store_code = '$loc6' 
									or store_code = '$loc7' or store_code = '$loc8' or store_code = '$loc9' or store_code = '$loc10' or store_code = '$loc11' 
									or store_code = '$loc12' or store_code = '$loc13' or store_code = '$loc14' or store_code = '$loc15' or store_code = '$loc16' 
									or store_code = '$loc17' or store_code = '$loc18' or store_code = '$loc19' or store_code = '$loc20' or store_code = '$loc21' 
									or store_code = '$loc22' or store_code = '$loc23' or store_code = '$loc24' or store_code = '$loc25' or store_code = '$loc26' 
									or store_code = '$loc27' or store_code = '$loc28' or store_code = '$loc29' or store_code = '$loc30' or store_code = '$loc31' 
									or store_code = '$loc32' or store_code = '$loc33' or store_code = '$loc34' or store_code = '$loc35' or store_code = '$loc36'
									or store_code = '$loc37') 
									and (division <> '1500' and division <> '8000' and division <> '8500' and division <> '9000')
								 GROUP BY attrib1, attrib2 
								 ORDER BY attrib1
								});	
	$sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$group_code = $s->{attrib1};
		$group_desc = $s->{attrib2};
				
		$sls3 = $dbh_csv->prepare (qq{SELECT division, div_name, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
									 FROM $table 
									 WHERE merch_group_code_rev = '$merch_group_code' AND attrib1 = '$group_code' and 
										(store_code = '$loc1' or store_code = '$loc2' or store_code = '$loc3' or store_code = '$loc4' or store_code = '$loc5' 
										or store_code = '$loc6' or store_code = '$loc7' or store_code = '$loc8' or store_code = '$loc9' or store_code = '$loc10' 
										or store_code = '$loc11' or store_code = '$loc12' or store_code = '$loc13' or store_code = '$loc14' or store_code = '$loc15' 
										or store_code = '$loc16' or store_code = '$loc17' or store_code = '$loc18' or store_code = '$loc19' or store_code = '$loc20' 
										or store_code = '$loc21' or store_code = '$loc22' or store_code = '$loc23' or store_code = '$loc24' or store_code = '$loc25' 
										or store_code = '$loc26' or store_code = '$loc27' or store_code = '$loc28' or store_code = '$loc29' or store_code = '$loc30' 
										or store_code = '$loc31' or store_code = '$loc32' or store_code = '$loc33' or store_code = '$loc34' or store_code = '$loc35' 
										or store_code = '$loc36' or store_code = '$loc37')
										and (division <> '1500' and division <> '8000' and division <> '8500' and division <> '9000')
									 GROUP BY division, div_name 
									 ORDER BY division
									});
		$sls3->execute();
		
		$grp_counter = $a;
		while(my $s = $sls3->fetchrow_hashref()){
			$division = $s->{division};
			$division_desc = $s->{div_name};
		
			if($stopper eq 0){
				$worksheet->merge_range( $a, 4, $a, 6, $division_desc, $desc );}
			
			$worksheet->write($a,$col, $s->{tot_items},$border1); 
			$worksheet->write($a,$col+1, $s->{rep_items},$border1);
				if ($s->{tot_items} <= 0){
					$worksheet->write($a,$col+2, "",$subt); 				}
				else{
					$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt); 				}
			
			$counter = 0; #RESET dept_counter	
			$a++; #INCREMENT VARIABLE a
		}

		if($group_code ne 'JW'){
			$grp_tot_items += $s->{tot_items};
			$grp_rep_items += $s->{rep_items};
		}
		
		$worksheet->write($a,$col, $s->{tot_items},$bodyNum); 
		$worksheet->write($a,$col+1, $s->{rep_items},$bodyNum);
		if ($s->{tot_items} <= 0){
			$worksheet->write($a,$col+2, "",$bodyPct);
		}
		else{
			$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$bodyPct);
		}
		
		if($stopper eq 0){
			$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );
		}
		
		$a++; #INCREMENT VARIABLE a
	}
	
	$total_tot_items += $s->{tot_items};
	$total_rep_items += $s->{rep_items};
	
	if ($merch_group_code eq 'DS'){
		############DEPT STORE WO JEWELRY ###############
		$worksheet->write($a,$col, $grp_tot_items,$headNumber); 
		$worksheet->write($a,$col+1, $grp_rep_items,$headNumber);
		if ($grp_tot_items <= 0){
			$worksheet->write($a,$col+2, "",$headPct);
		}
		else{
			$worksheet->write($a,$col+2, '=IF(ISERROR('. $grp_rep_items/$grp_tot_items .'),"",('.$grp_rep_items/$grp_tot_items. '))',$headPct);
		}
		
		if($stopper eq 0){
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store w/o Jewelry', $headN );
		}
		$a += 1;
		
		if($stopper eq 0){
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
			$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
		}
	}
	elsif($merch_group_code eq 'SU' and $stopper eq 0 ){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
	}
	
	$worksheet->write($a,$col, $s->{tot_items},$headNumber); 
	$worksheet->write($a,$col+1, $s->{rep_items},$headNumber);
	if ($s->{tot_items} <= 0){
		$worksheet->write($a,$col+2, "",$headPct);
	}
	else{
		$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$headPct);
	}
	
	$a++; #INCREMENT VARIABLE a
}

	$worksheet->write($a,$col, $total_tot_items,$headNumber); 
	$worksheet->write($a,$col+1, $total_rep_items,$headNumber);
	if ($total_tot_items <= 0){
		$worksheet->write($a,$col+2, "",$headPct);
	}
	else{
		$worksheet->write($a,$col+2, '=IF(ISERROR('. $total_rep_items/$total_tot_items .'),"",('.$total_rep_items/$total_tot_items. '))',$headPct);
	}

if($stopper eq 0){		
	$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );
}

$worksheet->merge_range( $loc_pt, $col, $loc_pt, $col+2, $region, $subhead );

$sls1->finish();
$sls2->finish();
$sls3->finish();
#$sls4->finish();

$len = $a, $counter = 0;

}

sub query_dept_store {

$table = 'instock_nonrep_v1.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls1 = $dbh_csv->prepare (qq{SELECT store_code, store_name, merch_group_code_rev, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
								 FROM $table 
								 WHERE store_code IN ('$loc') and (division <> '1500' and division <> '8000' and division <> '8500' and division <> '9000')
								 GROUP BY store_code, store_name, merch_group_code_rev 
								 ORDER BY store_code, merch_group_code_rev
							 });								 
$sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$merch_group_code = $s->{merch_group_code_rev};
	$loc_code = $s->{store_code};
	$loc_desc = $s->{store_name};
	
	$sls2 = $dbh_csv->prepare (qq{SELECT attrib1, attrib2, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
								 FROM $table 
								 WHERE merch_group_code_rev = '$merch_group_code' AND store_code IN ('$loc') and (division <> '1500' and division <> '8000' and division <> '8500' and division <> '9000')
								 GROUP BY attrib1, attrib2 
								 ORDER BY attrib1
								});	
	$sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$group_code = $s->{attrib1};
		$group_desc = $s->{attrib2};
				
		$sls3 = $dbh_csv->prepare (qq{SELECT division, div_name, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
									 FROM $table 
									 WHERE merch_group_code_rev = '$merch_group_code' AND attrib1 = '$group_code' and store_code IN ('$loc') 
											and (division <> '1500' and division <> '8000' and division <> '8500' and division <> '9000')
									 GROUP BY division, div_name 
									 ORDER BY division
									});
		$sls3->execute();
		
		$grp_counter = $a;
		while(my $s = $sls3->fetchrow_hashref()){
			$division = $s->{division};
			$division_desc = $s->{div_name};
			
			$sls4 = $dbh_csv->prepare (qq{SELECT group_no, group_name, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
										 FROM $table 
										 WHERE (merch_group_code_rev = '$merch_group_code') AND attrib1 = '$group_code' and division = '$division' and store_code IN ('$loc') 
												and (division <> '1500' and division <> '8000' and division <> '8500' and division <> '9000')
										 GROUP BY group_no, group_name 
										 ORDER BY group_no
										 });
			$sls4->execute();
			
			while(my $s = $sls4->fetchrow_hashref()){		
				$worksheet->write($a,5, $s->{group_no},$code);
				$worksheet->write($a,6, $s->{group_name},$desc);
				$worksheet->write($a,$col, $s->{tot_items},$border1); 
				$worksheet->write($a,$col+1, $s->{rep_items},$border1);
				if ($s->{tot_items} <= 0){
					$worksheet->write($a,$col+2, "",$subt);
				}
				else{
					$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt);
				}
				
				$a++;
				$counter++;
		
			}
			
			&calc; #division subtotal
			if($stopper eq 0){
				$worksheet->merge_range( $a, 5, $a, 6, 'Subtotal', $bodyN );
				$worksheet->merge_range( $a-$counter, 4, $a, 4, $division_desc, $border2 );
			}
			
			$counter = 0; #RESET dept_counter	
			$a++; #INCREMENT VARIABLE a
		}

		if($group_code ne 'JW'){
			$grp_tot_items += $s->{tot_items};
			$grp_rep_items += $s->{rep_items};
		}
		
		$worksheet->write($a,$col, $s->{tot_items},$bodyNum); 
		$worksheet->write($a,$col+1, $s->{rep_items},$bodyNum);
		if ($s->{tot_items} <= 0){
			$worksheet->write($a,$col+2, "",$bodyPct);
		}
		else{
			$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$bodyPct);
		}
		
		if($stopper eq 0){
			$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );
			$worksheet->merge_range( $grp_counter, 3, $a, 3, $group_desc, $border2 );
		}
		
		$a++; #INCREMENT VARIABLE a
	}
	
	$total_tot_items += $s->{tot_items};
	$total_rep_items += $s->{rep_items};
	
	if ($merch_group_code eq 'DS'){
		############DEPT STORE WO JEWELRY ###############
		$worksheet->write($a,$col, $grp_tot_items,$headNumber); 
		$worksheet->write($a,$col+1, $grp_rep_items,$headNumber);
		if ($grp_tot_items <= 0){
			$worksheet->write($a,$col+2, "",$headPct);
		}
		else{
			$worksheet->write($a,$col+2, '=IF(ISERROR('. $grp_rep_items/$grp_tot_items .'),"",('.$grp_rep_items/$grp_tot_items. '))',$headPct);
		}
		
		if($stopper eq 0){
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store w/o Jewelry', $headN );
		}
		$a += 1;
		
		if($stopper eq 0){
			$worksheet->merge_range( $a, 3, $a, 6, 'Total Department Store', $headN );
			$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'DEPARTMENT STORE', $border2 );
		}
	}
	elsif($merch_group_code eq 'SU' and $stopper eq 0 ){
		$worksheet->merge_range( $a, 3, $a, 6, 'Total Supermarket', $headN );
		$worksheet->merge_range( $mgc_counter, 2, $a, 2, 'SUPERMARKET', $border2 );
	}
	
	$worksheet->write($a,$col, $s->{tot_items},$headNumber); 
	$worksheet->write($a,$col+1, $s->{rep_items},$headNumber);
	if ($s->{tot_items} <= 0){
		$worksheet->write($a,$col+2, "",$headPct);
	}
	else{
		$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$headPct);
	}
	
	$a++; #INCREMENT VARIABLE a
}

	$worksheet->write($a,$col, $total_tot_items,$headNumber); 
	$worksheet->write($a,$col+1, $total_rep_items,$headNumber);
	if ($total_tot_items <= 0){
		$worksheet->write($a,$col+2, "",$headPct);
	}
	else{
		$worksheet->write($a,$col+2, '=IF(ISERROR('. $total_rep_items/$total_tot_items .'),"",('.$total_rep_items/$total_tot_items. '))',$headPct);
	}

if($stopper eq 0){		
	$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN );
}

$worksheet->merge_range( $loc_pt, $col, $loc_pt, $col+2, $loc_code .'-'. $loc_desc, $subhead );

$sls1->finish();
$sls2->finish();
$sls3->finish();
$sls4->finish();

$len = $a, $counter = 0;

}


sub calc { #CALCULATION FOR EACH DIVISION

foreach my $c( $col..$col+2 ){
	my $sum = '=SUM('. xl_rowcol_to_cell( $a-$counter, $c ). ':' . xl_rowcol_to_cell( $a-1, $c ) . ')';
		$worksheet->write( $a, $c, $sum, $bodyNum );
		
		if ($c eq $col+2){
			my $pct = '=IFERROR('. xl_rowcol_to_cell( $a, $c-1 ). '/' . xl_rowcol_to_cell( $a, $c-2 ) .',)';
				$worksheet->write( $a, $c, $pct, $body );
		}
}

}

sub calc8 { 

if($type_test eq 3 or $type_test eq 4){
	foreach my $col( 7, 10, 13 ){
		if($s1f2_row ne 0){
			my $sum = '=(SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . '))-'. xl_rowcol_to_cell( $s1f2_row, $col );
				$worksheet->write( $a, $col, $sum, $bodyNum );
			my $sum_LY = '=(SUM('. xl_rowcol_to_cell( $a-$counter, $col+1 ). ':' . xl_rowcol_to_cell( $a-1, $col+1 ) . '))-'. xl_rowcol_to_cell( $s1f2_row, $col+1 );
				$worksheet->write( $a, $col+1, $sum_LY, $bodyNum );
			
				if( '=SUM('. xl_rowcol_to_cell( $a-$counter, $col+1 ). ':' . xl_rowcol_to_cell( $a-1, $col+1 ) . ')' le 0){
					$worksheet->write( $a, $col+2, '', $bodyPct );
				}
				else{
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col+1 ). '/' . xl_rowcol_to_cell( $a, $col ) . ')';
					$worksheet->write( $a, $col+2, $pct2sls, $bodyPct );	
				}
				
		}
		else{
			my $sum = '=SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . ')';
				$worksheet->write( $a, $col, $sum, $bodyNum );
			my $sum_LY = '=SUM('. xl_rowcol_to_cell( $a-$counter, $col+1 ). ':' . xl_rowcol_to_cell( $a-1, $col+1 ) . ')';
				$worksheet->write( $a, $col+1, $sum_LY, $bodyNum );

				if( '=SUM('. xl_rowcol_to_cell( $a-$counter, $col+1 ). ':' . xl_rowcol_to_cell( $a-1, $col+1 ) . ')' le 0){
					$worksheet->write( $a, $col+2, '', $bodyPct );
					print "test";
				}
				else{
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col+1 ). '/' . xl_rowcol_to_cell( $a, $col ) . ')';
					$worksheet->write( $a, $col+2, $pct2sls, $bodyPct );
				}
		}		
	}
}

else{
	foreach my $col( 7, 8, 10, 12, 13, 15 ){
		if($s1f2_row ne 0){
			my $sum = '=(SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . '))-'. xl_rowcol_to_cell( $s1f2_row, $col );
			$worksheet->write( $a, $col, $sum, $bodyNum );
			
			if ($col eq 8 or $col eq 13){
				if( xl_rowcol_to_cell( $a, $col ) eq 0){
					$worksheet->write( $a, $col+1, '', $bodyPct );
				}
				else{
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
				}
			}
		}
		else{
			my $sum = '=SUM('. xl_rowcol_to_cell( $a-$counter, $col ). ':' . xl_rowcol_to_cell( $a-1, $col ) . ')';
				$worksheet->write( $a, $col, $sum, $bodyNum );	
				
			if ($col eq 8 or $col eq 13){
				if( xl_rowcol_to_cell( $a, $col ) eq 0){
					$worksheet->write( $a, $col+1, '', $bodyPct );
				}
				else{
					my $pct2sls = '=('. xl_rowcol_to_cell( $a, $col-1 ). '/' . xl_rowcol_to_cell( $a, $col ) . '-1' . ')';
					$worksheet->write( $a, $col+1, $pct2sls, $bodyPct );
				}
			}
		}		
	}
}
}

sub calc_tot_region { #TOTAL CALCULATION

$worksheet->merge_range( 4, $col, 4, $col+2, $total_label, $subhead );
$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

if($end eq 1){
	foreach my $c( 6..$len ){
		my $sumCount = '=SUMIFS('.xl_rowcol_to_cell($c,7).':'.xl_rowcol_to_cell($c,$col-2).','. xl_rowcol_to_cell(5,7).':'. xl_rowcol_to_cell(5,$col-2).',"COUNT")';
			$worksheet->write( $c, $col, $sumCount, $headNumber );
		my $sumStock = '=SUMIFS('.xl_rowcol_to_cell($c,7).':'.xl_rowcol_to_cell($c,$col-2).','. xl_rowcol_to_cell(5,7).':'. xl_rowcol_to_cell(5,$col-2).',"IN STOCK")';
			$worksheet->write( $c, $col+1, $sumStock, $headNumber );		
		my $pct = '=IFERROR('. xl_rowcol_to_cell( $c, $col+1 ). '/' . xl_rowcol_to_cell( $c, $col ).', )' ;
			$worksheet->write( $c, $col+2, $pct, $headN );
	}
}
elsif($end eq 0){
	foreach my $c( 6..$len ){
		my $sumCount = '=SUMIFS('.xl_rowcol_to_cell($c,$vis).':'.xl_rowcol_to_cell($c,$col-1).','. xl_rowcol_to_cell(5,$vis).':'. xl_rowcol_to_cell(5,$col-1).',"COUNT")';
			$worksheet->write( $c, $col, $sumCount, $headNumber );
		my $sumStock = '=SUMIFS('.xl_rowcol_to_cell($c,$vis).':'.xl_rowcol_to_cell($c,$col-1).','. xl_rowcol_to_cell(5,$vis).':'. xl_rowcol_to_cell(5,$col-1).',"IN STOCK")';
			$worksheet->write( $c, $col+1, $sumStock, $headNumber );		
		my $pct = '=IFERROR('. xl_rowcol_to_cell( $c, $col+1 ). '/' . xl_rowcol_to_cell( $c, $col ).', )' ;
			$worksheet->write( $c, $col+2, $pct, $headN );
	}
}

foreach my $i ("T COUNT ", "T IN STOCK", "%") {
	$worksheet->write(5, $col++, $i, $subhead);
}

$vis = $col;

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
 open my $fh, ">", "instock_nonrep_v1.csv" or die "instock_nonrep_v1.csv: $!";

#version 1.21 addl filter, last received within 3mo for SU/last received within 6mo for DS
$test = qq{
SELECT 
CASE WHEN TO_CHAR(SGD.STORE) = '4002' THEN '2001W' ELSE TO_CHAR(SGD.STORE) END AS STORE_CODE,
CASE WHEN SGD.STORE IN ('2012', '2013', '3009', '4004', '3010', '3011', '3012') THEN 'SU' || SGD.STORE	 WHEN SGD.STORE IN ('4002') THEN 'SU2001W'     WHEN SGD.STORE = '2223' THEN 'DS' || SGD.STORE	 ELSE SGD.MERCH_GROUP_CODE || SGD.STORE END AS STORE,
SGD.STORE_NAME STORE_NAME, 
SGD.MERCH_GROUP_CODE, 
CASE WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 9000) THEN 'DS'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8500) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO = 8040) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO != 8040) THEN 'DS'ELSE SGD.MERCH_GROUP_CODE END AS MERCH_GROUP_CODE_REV,
SGD.MERCH_GROUP_DESC, 
SGD.ATTRIB1, SGD.ATTRIB2, 
SGD.DIVISION, SGD.DIV_NAME, 
SGD.GROUP_NO GROUP_NO, SGD.GROUP_NAME GROUP_NAME,
SUM(STK.TOT_NON_REP_ITEMS) TOT_REPL_ITEMS, 
SUM(STK.NON_REP_WITH_SOH) REPL_WITH_SOH
FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		FROM
			(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (3010,6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
			(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
				FROM DEPS D 
				  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
				  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
				  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
		GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
	)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(NON_REPL_TAG) TOT_NON_REP_ITEMS, COUNT(STOCK_TAG) NON_REP_WITH_SOH
	FROM(
		SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, SOH.ITEM, SOH.LOC LOCATION, LOCATIONS.LOC_NAME, 
		  CASE WHEN REPL.ITEM IS NULL THEN 'Y' END AS NON_REPL_TAG, SOH.STOCK_ON_HAND, 
		  CASE WHEN (REPL.ITEM IS NULL AND SOH.STOCK_ON_HAND <= 0) THEN NULL
			   WHEN (REPL.ITEM IS NOT NULL AND SOH.STOCK_ON_HAND > 0) THEN NULL
			   ELSE 'Y' END AS STOCK_TAG 
		FROM ITEM_LOC_SOH SOH
			LEFT JOIN ITEM_LOC LOC ON SOH.ITEM=LOC.ITEM AND SOH.LOC=LOC.LOC
			LEFT JOIN REPL_ITEM_LOC REPL ON SOH.ITEM=REPL.ITEM AND SOH.LOC=REPL.LOCATION AND REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
			LEFT JOIN ITEM_MASTER MST ON SOH.ITEM=MST.ITEM 
			LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
			LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
			LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON SOH.LOC=LOCATIONS.LOC
		WHERE LOC.STATUS = 'A' AND DEPS.PURCHASE_TYPE = 0
		)TBL
	GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK
	ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO
GROUP BY 
CASE WHEN TO_CHAR(SGD.STORE) = '4002' THEN '2001W' ELSE TO_CHAR(SGD.STORE) END,
CASE WHEN SGD.STORE IN ('2012', '2013', '3009', '4004', '3010', '3011', '3012') THEN 'SU' || SGD.STORE 	 WHEN SGD.STORE IN ('4002') THEN 'SU2001W'     WHEN SGD.STORE = '2223' THEN 'DS' || SGD.STORE	 ELSE SGD.MERCH_GROUP_CODE || SGD.STORE END,	 SGD.STORE_NAME, 
SGD.MERCH_GROUP_CODE, 
CASE WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 9000) THEN 'DS'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8500) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO = 8040) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO != 8040) THEN 'DS'ELSE SGD.MERCH_GROUP_CODE END,
SGD.MERCH_GROUP_DESC, 
SGD.ATTRIB1, SGD.ATTRIB2, 
SGD.DIVISION, SGD.DIV_NAME, 
SGD.GROUP_NO, SGD.GROUP_NAME
ORDER BY 1, 3, 5, 7, 9
}; 
 
my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_lc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "instock_nonrep_v1.csv: $!";
 
$dbh->disconnect;

}

sub zip_file {

# Create a Zip file
my $zip = Archive::Zip->new();
   
# add a directory
#my $dir_member = $zip->addDirectory( 'C:\\Perl\\BI_Reports\\RMS Reports\\' );
   
# create/add a file from a string with compression
# my $string_member = $zip->addString( 'This is a system generated report.', 'about_file.txt' );
# $string_member->desiredCompressionMethod( COMPRESSION_DEFLATED );
   
# add the file to zip
my $file_member = $zip->addFile( 'instock_raw_file_v12.csv' );
   
# save the zip file
unless ( $zip->writeToFileNamed('instock_raw_file_v12.rar') == AZ_OK ) {
	die 'write error';
}
   
}


sub mail {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file, $attachment_file_2 ) = @ARGV;

$to = 'cherry.gulloy@metroretail.com.ph, janice.bedrijo@metroretail.com.ph, jerson.roma@metroretail.com.ph, bermon.alcantara@metroretail.com.ph, eljie.laquinon@metroretail.com.ph, nilynn.yosores@metroretail.com.ph, ryanneil.dupay@metroretail.com.ph, anafatima.mancho@metroretail.com.ph ';

$cc = 'annalyn.conde@metroretail.com.ph, lea.gonzaga@metroretail.com.ph,eric.molina@metroretail.com.ph';
#  $cc = 'lea.gonzaga@metroretail.com.ph,eric.molina@metroretail.com.ph';

$from = 'Report Mailer<report.mailer@metroretail.com.ph>';

$subject = 'Non Replenishment In-stock';

$msgbody_file = 'message.txt';

$attachment_file = "Non Replenishment In-stock v1.xlsx";

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
#primary recipients
sub mail1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file, $attachment_file_2 ) = @ARGV;

$to = ' arthur.emmanuel@metroretail.com.ph,fili.mercado@metroretail.com.ph,emily.silverio@metroretail.com.ph,ronald.dizon@metroretail.com.ph,chit.lazaro@metroretail.com.ph,jocelyn.sarmiento@metroretail.com.ph,charisse.mancao@metroretail.com.ph,cindy.yu@metroretail.com.ph,cresilda.dehayco@metroretail.com.ph,evan.inocencio@metroretail.com.ph,fe.botero@metroretail.com.ph,jonrel.nacor@metroretail.com.ph,junah.oliveron@metroretail.com.ph,lyn.cabatuan@metroretail.com.ph,zenda.mangabon@metroretail.com.ph,joyce.mirabueno@metroretail.com.ph,mariegrace.ong@metroretail.com.ph,cherry.gulloy@metroretail.com.ph,janice.bedrijo@metroretail.com.ph,jerson.roma@metroretail.com.ph,bermon.alcantara@metroretail.com.ph,nilynn.yosores@metroretail.com.ph,anafatima.mancho@metroretail.com.ph,emily.silverio@metroretail.com.ph,leslie.chipeco@metroretail.com.ph,karan.malani@metroretail.com.ph ';
$cc = 'luz.bitang@metroretail.com.ph,rex.cabanilla@metroretail.com.ph, annalyn.conde@metroretail.com.ph';
$bcc = 'lea.gonzaga@metroretail.com.ph,eric.molina@metrogaisano.com';

# $to = 'kent.mamalias@metroretail.com.ph';  		

$subject = 'Replenishment In-stock';
$msgbody_file = 'message.txt';

$attachment_file = "Replenishment In-stock v1.21.xlsx";
#$attachment_file_raw = "instock_raw_file_v12.rar";

my $msgbody = read_file( $msgbody_file );

my $attachment_data = encode_base64( read_file( $attachment_file, 1 ));
#my $attachment_data_2 = encode_base64( read_file( $attachment_file_2, 1 ));

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
#area/store managers first part
sub mail2 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = 'manuel.degamo@metroretail.com.ph, ace.olalia@metroretail.com.ph, alma.espino@metroretail.com.ph, angeli_christi.ladot@metroretail.com.ph, angelito.dublin@metroretail.com.ph, arlene.yanson@metroretail.com.ph, augosto.daria@metroretail.com.ph,  flor.bolante@metroretail.com.ph, teena.velasco@metroretail.com.ph, cristy.sy@metroretail.com.ph, diana.almagro@metroretail.com.ph, edgardo.lim@metroretail.com.ph, edris.tarrobal@metroretail.com.ph, fidela.villamor@metroretail.com.ph, genaro.felisilda@metroretail.com.ph, genevive.quinones@metroretail.com.ph, glenda.navares@metroretail.com.ph, joefrey.camu@metroretail.com.ph, jonalyn.diaz@metroretail.com.ph, opcplanning@metroretail.com.ph,mopcplanning.foodretail@metroretail.com.ph,eric.molina@metroretail.com.ph ';
$cc = 'lea.gonzaga@metroretail.com.ph,eric.molina@metroretail.com.ph';
		
$subject = 'Replenishment In-stock';
$msgbody_file = 'message.txt';

$attachment_file = "Replenishment In-stock v1.21.xlsx";

my $msgbody = read_file( $msgbody_file );

my $attachment_data = encode_base64( read_file( $attachment_file, 1 ));

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
#area/store managers second part
sub mail3 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = ' josemarie.graciadas@metroretail.com.ph, jovany.polancos@metroretail.com.ph, judy.gilo@metroretail.com.ph, julie.montano@metroretail.com.ph, kathlene.procianos@metroretail.com.ph, limuel.ulanday@metroretail.com.ph, cristina.de_asis@metroretail.com.ph, mariajoana.cruz@metroretail.com.ph, may.sasedor@metroretail.com.ph, michelle.calsada@metroretail.com.ph, policarpo.mission@metroretail.com.ph, rex.refuerzo@metroretail.com.ph, ricky.tulda@metroretail.com.ph, ronald.dizon@metroretail.com.ph, roselle.agbayani@metroretail.com.ph, rowena.tangoan@metroretail.com.ph, roy.igot@metroretail.com.ph, tessie.cabanero@metroretail.com.ph, victoria.ferolino@metroretail.com.ph, wendel.gallo@metroretail.com.ph, juanjose.sibal@metroretail.com.ph, julie.montano@metroretail.com.ph ';
$cc = 'lea.gonzaga@metroretail.com.ph,eric.molina@metroretail.com.ph';

		
$subject = 'Replenishment In-stock';
$msgbody_file = 'message.txt';

$attachment_file = "Replenishment In-stock v1.21.xlsx";

my $msgbody = read_file( $msgbody_file );

my $attachment_data = encode_base64( read_file( $attachment_file, 1 ));

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








