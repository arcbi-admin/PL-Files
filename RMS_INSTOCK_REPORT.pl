use DBI;
use DBD::Oracle qw(:ora_types);
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use Text::CSV_XS;
use Win32::Job;
use Getopt::Long;
use IO::File;
use MIME::QuotedPrint;
use MIME::Base64;
use Mail::Sendmail;
use Date::Calc qw( Today Add_Delta_Days Month_to_Text);

($year,$month,$day) = Today();
$month_to_text = Month_to_Text($month);

 # my $mms_job = Win32::Job->new;
	# $mms_job->spawn( "cmd" , q{cmd /C "java ecp_markdown pause"});
	# $mms_job->run(1500);

my $workbook = Excel::Writer::XLSX->new("INSTOCK_REPORT.xlsx");
my $bold = $workbook->add_format( bold => 1, size => 14 );
my $bold1 = $workbook->add_format( bold => 1, size => 11 );
my $bold2 = $workbook->add_format( size => 11 );
my $border1 = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3 );
my $border2 = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', text_wrap =>1, size => 10, shrink => 1 );
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
my $headNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3, bg_color => $abo, bold => 1 );
my $headNumber = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 3, bg_color => $abo, bold => 1 );
my $head = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
my $subhead = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, bg_color => $ponkan, bold => 1 );
my $bodyN = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
my $bodyNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 3,  bold => 1);
my $body = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
my $subt = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 9);
my $down = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => 9, bg_color => $pula );

printf "IN STOCK REPORT \n";

&generate_csv;

&new_sheet($sheet = "ByDept-Store", $comment = "STORES");		
&s1;
&s2;
&s3;
&s8;
&s10;
&s11;
&f2;
&f3;
&f6;
&h1;
&h3;
&h5;
&h10;
&tg1;
&tg2;
&tg3;
&tg4;
&tg5;
&tg6;
&h4;
&h8;
&calc_tot_region($total_label = 'TOTAL VISAYAS', $end = 0);

&s4;
&s5;
&s6;
&s7;
&s9;
&s12;
&s13;
&tg7;
&f4;
&h2;
&calc_tot_region($total_label = 'TOTAL LUZON', $end = 0);
&calc_tot_region($total_label = 'TOTAL', $end = 1);

&new_sheet($sheet = "ByDept-Whse", $comment = "WAREHOUSE");
&w80001;
&w80011;
&w80031;
&w80041;
&w80051;
&w80061;
&calc_tot_region($total_label = 'TOTAL WHSE', $end = 0);

$workbook->close();
$dbh_csv->disconnect;

#&mail;	
&mail1;	
&mail2;	
&mail3;	
&mail4;	
&mail5;	
exit;
 
#================================= FUNCTIONS ==================================#

sub s1 { $store = "S1"; $loc = "'2001'"; $vis = 7; $col = 7; $test = 1; &heading;	&call_div;	}
sub s2 { $store = "S2"; $loc = "'2002'"; &call_div;	}
sub s3 { $store = "S3"; $loc = "'2003'"; &call_div;	}
sub s4 { $store = "S4"; $loc = "'2004'"; &call_div;	}
sub s5 { $store = "S5"; $loc = "'2005'"; &call_div;	}
sub s6 { $store = "S6"; $loc = "'2006'"; &call_div;	}
sub s7 { $store = "S7"; $loc = "'2007'"; &call_div;	}
sub s8 { $store = "S8"; $loc = "'2008'"; &call_div;	}
sub s9 { $store = "S9"; $loc = "'2009'"; &call_div;	}
sub s10 { $store = "S10"; $loc = "'2010'"; &call_div;	}
sub s11 { $store = "S11"; $loc = "'2011'"; &call_div;	}
sub s12 { $store = "S12"; $loc = "'2012'"; &call_div;	}
sub s13 { $store = "S13"; $loc = "'2013'"; &call_div;	}
sub tg1 { $store = "TG1"; $loc = "'3001'"; &call_div;	}
sub tg2 { $store = "TG2"; $loc = "'3002'"; &call_div;	}
sub tg3 { $store = "TG3"; $loc = "'3003'"; &call_div;	}
sub tg4 { $store = "TG4"; $loc = "'3004'"; &call_div;	}
sub tg5 { $store = "TG5"; $loc = "'3005'"; &call_div;	}
sub tg6 { $store = "TG6"; $loc = "'3006'"; &call_div;	}
sub tg7 { $store = "TG7"; $loc = "'3007'"; &call_div;	}
sub f2 { $store = "F2"; $loc = "'4002'"; &call_div;	}
sub f3 { $store = "F3"; $loc = "'4003'"; &call_div;	}
sub f4 { $store = "F4"; $loc = "'4004'"; &call_div;	}
sub f6 { $store = "F6"; $loc = "'3009'"; &call_div;	}
sub h1 { $store = "H1"; $loc = "'6001'"; &call_div;	}
sub h2 { $store = "H2"; $loc = "'6002'"; &call_div;	}
sub h3 { $store = "H3"; $loc = "'6003'"; &call_div;	}
sub h4 { $store = "H4"; $loc = "'6004'"; &call_div;	}
sub h5 { $store = "H5"; $loc = "'6005'"; &call_div;	}
sub h8 { $store = "H8"; $loc = "'6008'"; &call_div;	}
sub h10 { $store = "H10"; $loc = "'6010'"; &call_div;	}
sub w80001 { $store = "Central WH"; $loc = "'80001'"; $vis = 7; $col = 7; $test = 1; &heading;	&call_div;	}
sub w80011 { $store = "J-KING"; $loc = "'80011'"; &call_div;	}
sub w80031 { $store = "NFA"; $loc = "'80031'"; &call_div;	}
sub w80041 { $store = "Procter"; $loc = "'80041'"; &call_div;	}
sub w80051 { $store = "Pagsabungan"; $loc = "'80051'"; &call_div;	}
sub w80061 { $store = "Schenker"; $loc = "'80061'"; &call_div;	}


sub call_div {

$a += 6, $ds = 0, $e += 6, $count=6, $counter=0;

$worksheet->merge_range( $a-2, $col, $a-2, $col+2, $store, $subhead );
$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&div1;
&div2;
&div3;
&div4;
&div5;
&div6;
&calc($code = "Department Store", $ds = $a); # Department Store Total
&div8;
&div9;
&div10;
&div11;
&calc($code = "Supermarket"); # Supermarket Total
&calc7; #TOTAL

$test = 0, $a = 0, $ds = 0, $e = 0, $counter=0; #RE INITIALIZE VARIABLES
$col += 3;

}

sub div1 {

$div_name = "Appliance and Home Improvement";
$g1 = '2530'; $g2 = '2560'; $g3 = '2570'; $g4 = '2580'; $g5 = '2590'; $g6 = '0000'; $g7 = '0000'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 
&query_markdown;

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0; #ADD 1 TO VARIABLES A RESET COUNTER
}

sub div2 {

$div_name = "Miscellaneous";
$g1 = '6540'; $g2 = '5510'; $g3 = '6550'; $g4 = '6510'; $g5 = '6570'; $g6 = '0000'; $g7 = '0000'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 
&query_markdown;

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0;
}

sub div3 {
$div_name = "Home Furnishings";
$g1 = '3020'; $g2 = '3010'; $g3 = '3040'; $g4 = '3050'; $g5 = '6530'; $g6 = '0000'; $g7 = '0000'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 
&query_markdown;

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0;
}

sub div4 {
$div_name = "Men\'s Lines";
$g1 = '3550'; $g2 = '5020'; $g3 = '7050'; $g4 = '0000'; $g5 = '0000'; $g6 = '0000'; $g7 = '0000'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 	
&query_markdown;

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0;
}

sub div5 {
$div_name = "Ladies\' Lines";
$g1 = '4540'; $g2 = '4520'; $g3 = '7030'; $g4 = '7040'; $g5 = '0000'; $g6 = '0000'; $g7 = '0000'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 
&query_markdown;

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0;
}

sub div6 {

$div_name = "Children\'s Lines";
$g1 = '3520'; $g2 = '3530'; $g3 = '3510'; $g4 = '5010'; $g5 = '7010'; $g6 = '7020'; $g7 = '3560'; $g8 = '3570'; $g9 = '3580'; $g10 = '3590'; 
&query_markdown; 

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0;
}

sub div8 {
$div_name = "FOOD";
$g1 = '1010'; $g2 = '1020'; $g3 = '1030'; $g4 = '1040'; $g5 = '1050'; $g6 = '1060'; $g7 = '2010'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 
&query_markdown;

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0;
}

sub div9 {
$div_name = "Non-Food";
$g1 = '5530'; $g2 = '5520'; $g3 = '0000'; $g4 = '0000'; $g5 = '0000'; $g6 = '0000'; $g7 = '0000'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 
&query_markdown;

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0;
}

sub div10 {
$div_name = "Pharmacy";
$g1 = '6010'; $g2 = '0000'; $g3 = '0000'; $g4 = '0000'; $g5 = '0000'; $g6 = '0000'; $g7 = '0000'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 
&query_markdown;

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0;
}

sub div11 {
$div_name = "FRESH";
$g1 = '2020'; $g2 = '2030'; $g3 = '2040'; $g4 = '2050'; $g5 = '2070'; $g6 = '0000'; $g7 = '0000'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 
&query_markdown;

if($test eq 1){
	if($a-$counter eq $a-1){
		$worksheet->write( $a-1, 3, $div_name, $border2 );
	}
	else{
		$worksheet->merge_range( $a-$counter, 3, $a-1, 3, $div_name, $border2 );
	}
}

$counter = 0;
}


sub new_sheet{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(85);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
$worksheet->set_print_scale( 100 );
$worksheet->set_margins( 0.05 );
$worksheet->conditional_formatting( 'F9:V100',  { type => 'cell', criteria => '<', value => 0, format => $down } );
$worksheet->set_column( 0, 2, 1 );
$worksheet->set_column( 3, 3, 17 );
$worksheet->set_column( 6, 6, 19 );
$worksheet->set_column( 1, 2, undef, undef, 1 );
$worksheet->set_column( 4, 5, undef, undef, 1 );
$worksheet->set_column( 34, 34, undef, undef, 1 );

}


sub heading {

$worksheet->write(0, 3, "IN STOCK " . $comment , $bold);
$worksheet->write(1, 3, $day . '-' . $month_to_text . '-' .$year, $bold2);
$worksheet->merge_range( 4, 3, 5, 3, 'DIV', $subhead );
$worksheet->merge_range( 4, 5, 5, 5, 'CODE', $subhead );
$worksheet->merge_range( 4, 6, 5, 6, 'DEPARTMENT', $subhead );

}

sub query_markdown {

$table = 'instock.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;

$sls = $dbh_csv->prepare (qq{SELECT group_no, group_name, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
								 FROM $table 
								 WHERE (group_no=$g1 or group_no=$g2 or group_no=$g3 or group_no=$g4 or group_no=$g5 or group_no=$g6 or group_no=$g7 or group_no=$g8 or group_no=$g9 or group_no=$g10)
										AND STORE IN ($loc) 
								 GROUP BY group_no, group_name ORDER BY group_no
								});

$sls->execute();
while(my $s = $sls->fetchrow_hashref()){
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
	$count++;
	$counter++;
}

$sls->finish();

}


sub calc { #CALCULATION FOR EACH DIVISION

foreach my $c( $col..$col+2 ){
	my $sum = '=SUM('. xl_rowcol_to_cell( $e, $c ). ':' . xl_rowcol_to_cell( $a-1, $c ) . ')';
		$worksheet->write( $a, $c, $sum, $bodyNum );
		
		if ($c eq $col+2){
			my $pct = '=IFERROR('. xl_rowcol_to_cell( $a, $c-1 ). '/' . xl_rowcol_to_cell( $a, $c-2 ) .',)';
				$worksheet->write( $a, $c, $pct, $body );
		}
}
if($test eq 1){
	$worksheet->merge_range( $a, 3, $a, 6, $code, $bodyN );
}

$a += 1; 
$e = $a;

}

sub calc7 { #TOTAL CALCULATION

foreach my $c( $col..$col+2 ){
	my $sumTY = '=SUM('.xl_rowcol_to_cell($ds,$c).','.xl_rowcol_to_cell($a-1,$c).')';
		$worksheet->write( $a, $c, $sumTY, $headNumber );
		
		if ($c eq $col+2){
			my $pct = '=IFERROR('. xl_rowcol_to_cell( $a, $c-1 ). '/' . xl_rowcol_to_cell( $a, $c-2 ) .',)';
				$worksheet->write( $a, $c, $pct, $head );
		}
}

if($test eq 1){
	$worksheet->merge_range( $a, 3, $a, 6, "TOTAL", $headN );
}
}

sub calc_tot_region { #TOTAL CALCULATION

$worksheet->merge_range( 4, $col, 4, $col+2, $total_label, $subhead );
$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 5 );

if($end eq 1){
	foreach my $c( 6..$count+2 ){
		my $sumCount = '=SUMIFS('.xl_rowcol_to_cell($c,7).':'.xl_rowcol_to_cell($c,$col-2).','. xl_rowcol_to_cell(5,7).':'. xl_rowcol_to_cell(5,$col-2).',"COUNT")';
			$worksheet->write( $c, $col, $sumCount, $headNumber );
		my $sumStock = '=SUMIFS('.xl_rowcol_to_cell($c,7).':'.xl_rowcol_to_cell($c,$col-2).','. xl_rowcol_to_cell(5,7).':'. xl_rowcol_to_cell(5,$col-2).',"IN STOCK")';
			$worksheet->write( $c, $col+1, $sumStock, $headNumber );		
		my $pct = '=IFERROR('. xl_rowcol_to_cell( $c, $col+1 ). '/' . xl_rowcol_to_cell( $c, $col ).', )' ;
			$worksheet->write( $c, $col+2, $pct, $headN );
	}
}
elsif($end eq 0){
	foreach my $c( 6..$count+2 ){
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
my $uid = 'kent';
my $pw = 'amer1c8';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh, ">", "instock.csv" or die "instock.csv: $!";

$test = qq{ 
SELECT CASE WHEN SGD.STORE = 2223 THEN 2013 ELSE SGD.STORE END AS STORE, SGD.STORE_NAME STORE_NAME, SGD.GROUP_NO GROUP_NO, SGD.GROUP_NAME GROUP_NAME, SUM(STK.TOT_REPL_ITEMS) TOT_REPL_ITEMS, SUM(STK.REPL_WITH_SOH) REPL_WITH_SOH FROM
(SELECT STORE, STORE_NAME, DIVISION, DIV_NAME, GROUP_NO, GROUP_NAME 
FROM
	(SELECT DISTINCT STORE, STORE_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH),   
	(SELECT G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME FROM DEPS D, GROUPS G, DIVISION I WHERE D.GROUP_NO = G.GROUP_NO AND G.DIVISION = I.DIVISION) 
GROUP BY STORE, STORE_NAME, DIVISION, DIV_NAME, GROUP_NO, GROUP_NAME
)SGD
LEFT JOIN
(SELECT GROUP_NO, DEPT, DEPT_NAME, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
FROM(
	SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
	  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG, REPL_B.W_ITEM AUTO_REP
	FROM REPL_ITEM_LOC REPL
		LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
		LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
		LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
		LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
		LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
		LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
		LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
	WHERE LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
	)TBL
--WHERE AUTO_REP = 'Y'
GROUP BY GROUP_NO, DEPT, DEPT_NAME, LOCATION, LOC_NAME)STK
ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO
GROUP BY SGD.STORE, SGD.STORE_NAME, SGD.GROUP_NO, SGD.GROUP_NAME
ORDER BY 1, 3
};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_lc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "instock.csv: $!";
 
$dbh->disconnect;

}

sub mail {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = ' cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, jerson.roma@metrogaisano.com, bermon.alcantara@metrogaisano.com, eljie.laquinon@metrogaisano.com, nilynn.yosores@metrogaisano.com, ryanneil.dupay@metrogaisano.com, anafatima.mancho@metrogaisano.com, opcplanning@metrogaisano.com, joy.dizon@metrogaisano.com, emily.silverio@metrogaisano.com, chit.lazaro@metrogaisano.com, jocelyn.sarmiento@metrogaisano.com, charisse.mancao@metrogaisano.com, cindy.yu@metrogaisano.com, cresilda.dehayco@metrogaisano.com, evan.inocencio@metrogaisano.com, fe.botero@metrogaisano.com, jonrel.nacor@metrogaisano.com, judith.tud@metrogaisano.com, junah.oliveron@metrogaisano.com, lyn.cabatuan@metrogaisano.com, zenda.mangabon@metrogaisano.com, joyce.mirabueno@metrogaisano.com, marygrace.ong@metrogaisano.com, rowena.tangoan@metrogaisano.com, genevive.quinones@metrogaisano.com, angeli_cristi.ladot@metrogaisano.com, edris.tarrobal@metrogaisano.com, charmaine.buenaventura@metrogaisano.com, augosto.daria@metrogaisano.com, noli.lee@metrogaisano.com, jovany.polancos@metrogaisano.com, alma.espino@metrogaisano.com, arlene.yanson@metrogaisano.com, josemarie.graciadas@metrogaisano.com, ronald.parragatos@metrogaisano.com, vivian.ablang@metrogaisano.com, emma.villoson@metrogaisano.com, rachel.riva@metrogaisano.com, roselle.agbayani@metrogaisano.com, michelle.calsada@metrogaisano.com, jonalyn.diaz@metrogaisano.com, joseph.landicho@metrogaisano.com, rex.refuerso@metrogaisano.com, al_rey.candia@metrogaisano.com, cheruvim.villaceran@metrogaisano.com, mae_flor.lauronal@metrogaisano.com, egan.pacquiao@metro.com.ph, evanguardia@metro.com.ph, rysalas@metro.com.ph, jdorimon@metro.com.ph, eestrera@metro.com.ph, efermo@metro.com.ph, irene.montemayor@metrogaisano.com, dinah.ramirez@metrogaisano.com, limuel.ulanday@metrogaisano.com, tricia.luntao@metrogaisano.com, jacqueline.cano@metrogaisano.com, joefrey.camu@metrogaisano.com, karan.malani@metrogaisano.com, lia.chipeco@metrogaisano.com, danice.tanael@metrogaisano.com, jennifer.moreno@metrogaisano.com, jessica.gaisano@metrogaisano.com, keith.poblete@metrogaisano.com, micah.alvarado@metrogaisano.com, michelle.someros@metrogaisano.com, peachy.aquino@metrogaisano.com, ryan.uson@metrogaisano.com, sheen.ducay@metrogaisano.com, annie.desuyo@metrogaisano.com, carmelita.intia@metrogaisano.com, consorcia.mullon@metrogaisano.com, contessa.fernandez@metrogaisano.com, delia.jakosalem@metrogaisano.com, editha.cabriles@metrogaisano.com, edna.prieto@metrogaisano.com, jecil.cumayas@metrogaisano.com, jennifer.yu@metrogaisano.com, lorena.madraga@metrogaisano.com, maryann.delarama@metrogaisano.com, maryjoy.montes@metrogaisano.com, mecelle.quimbo@metrogaisano.com, mirasol.barcoma@metrogaisano.com, nenita.cabigon@metrogaisano.com, teresita.manatad@metrogaisano.com, vilma.paner@metrogaisano.com, melinda.uy@metrogaisano.com, jgeniston@metro.com.ph, cartes@metro.com.ph, hcaberte@metro.com.ph, mvilla@metro.com.ph, mcabungcal@metro.com.ph, mcombinido@metro.com.ph, april.agapito@metrogaisano.com, jordan.mok@metrogaisano.com, marlita.portes@metrogaisano.com, advento.resma@metrogaisano.com, arlene.te@metrogaisano.com, armando.pitogo@metrogaisano.com, christine.lanohan@metrogaisano.com, delvie.pitogo@metrogaisano.com, emillie.ponsica@metrogaisano.com, marlit.ignacio@metrogaisano.com, mildred.quinones@metrogaisano.com, rosemarie.saravia@metrogaisano.com, rowena.conde@metrogaisano.com, tessie.baldezamo@metrogaisano.com, alain.reyes@metrogaisano.com, diana.almagro@metrogaisano.com, angelito.dublin@metrogaisano.com, arlene.yanson@metrogaisano.com, charm.buenaventura@metrogaisano.com, teena.velasco@metrogaisano.com, cristy.sy@metrogaisano.com, fidela.villamor@metrogaisano.com, glenda.navares@metrogaisano.com, jonalyn.diaz@metrogaisano.com, josemarie.graciadas@metrogaisano.com, judy.gilo@metrogaisano.com, cristina.de_asis@metrogaisano.com, mariajoana.cruz@metrogaisano.com, ricky.tulda@metrogaisano.com, roy.igot@metrogaisano.com, may.sasedor@metrogaisano.com, donna.fernando@metrogaisano.com, victoria.ferolino@metrogaisano.com, wendel.gallo@metrogaisano.com, ronald.dizon@metrogaisano.com, alexander.tejedor@metrogaisano.com ';
$bcc = 'kent.mamalias@metrogaisano.com, rex.cabanilla@metrogaisano.com, lea.gonzaga@metrogaisano.com, fili.mercado@metrogaisano.com, annalyn.conde@metrogaisano.com';
		
# $to = 'kent.mamalias@metrogaisano.com';
# $bcc = 'kent.mamalias@metrogaisano.com';
		
$subject = 'In Stock Report';
$msgbody_file = 'message.txt';

$attachment_file = "INSTOCK_REPORT.xlsx";

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
Content-Type: text/plain; charset="iso-8859-1"
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

sub mail1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = ' cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, jerson.roma@metrogaisano.com, bermon.alcantara@metrogaisano.com, eljie.laquinon@metrogaisano.com, nilynn.yosores@metrogaisano.com, ryanneil.dupay@metrogaisano.com, anafatima.mancho@metrogaisano.com, opcplanning@metrogaisano.com, joy.dizon@metrogaisano.com, emily.silverio@metrogaisano.com, chit.lazaro@metrogaisano.com, jocelyn.sarmiento@metrogaisano.com, charisse.mancao@metrogaisano.com, cindy.yu@metrogaisano.com, cresilda.dehayco@metrogaisano.com, evan.inocencio@metrogaisano.com, fe.botero@metrogaisano.com, jonrel.nacor@metrogaisano.com, judith.tud@metrogaisano.com, junah.oliveron@metrogaisano.com, lyn.cabatuan@metrogaisano.com, zenda.mangabon@metrogaisano.com, joyce.mirabueno@metrogaisano.com, marygrace.ong@metrogaisano.com, rowena.tangoan@metrogaisano.com, genevive.quinones@metrogaisano.com, angeli_cristi.ladot@metrogaisano.com, edris.tarrobal@metrogaisano.com, charmaine.buenaventura@metrogaisano.com, augosto.daria@metrogaisano.com, cj.jesena@metrogaisano.com, gerry.guanlao@metrogaisano.com ';
$bcc = 'kent.mamalias@metrogaisano.com, rex.cabanilla@metrogaisano.com, lea.gonzaga@metrogaisano.com, fili.mercado@metrogaisano.com, annalyn.conde@metrogaisano.com';
		
# $to = 'kent.mamalias@metrogaisano.com';
# $bcc = 'kent.mamalias@metrogaisano.com';
		
$subject = 'In Stock Report';
$msgbody_file = 'message.txt';

$attachment_file = "INSTOCK_REPORT.xlsx";

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
Content-Type: text/plain; charset="iso-8859-1"
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

sub mail2 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = ' noli.lee@metrogaisano.com, jovany.polancos@metrogaisano.com, alma.espino@metrogaisano.com, arlene.yanson@metrogaisano.com, josemarie.graciadas@metrogaisano.com, ronald.parragatos@metrogaisano.com, vivian.ablang@metrogaisano.com, emma.villoson@metrogaisano.com, rachel.riva@metrogaisano.com, roselle.agbayani@metrogaisano.com, michelle.calsada@metrogaisano.com, jonalyn.diaz@metrogaisano.com, joseph.landicho@metrogaisano.com, rex.refuerso@metrogaisano.com, al_rey.candia@metrogaisano.com, cheruvim.villaceran@metrogaisano.com, mae_flor.lauronal@metrogaisano.com, egan.pacquiao@metro.com.ph, evanguardia@metro.com.ph, rysalas@metro.com.ph, jdorimon@metro.com.ph, eestrera@metro.com.ph, efermo@metro.com.ph, irene.montemayor@metrogaisano.com, dinah.ramirez@metrogaisano.com, limuel.ulanday@metrogaisano.com, tricia.luntao@metrogaisano.com, jacqueline.cano@metrogaisano.com, joefrey.camu@metrogaisano.com, karan.malani@metrogaisano.com, lia.chipeco@metrogaisano.com, danice.tanael@metrogaisano.com';
$cc = 'cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, jerson.roma@metrogaisano.com, bermon.alcantara@metrogaisano.com, eljie.laquinon@metrogaisano.com, nilynn.yosores@metrogaisano.com, ryanneil.dupay@metrogaisano.com, anafatima.mancho@metrogaisano.com';
$bcc = 'kent.mamalias@metrogaisano.com';
		
# $to = 'kent.mamalias@metrogaisano.com';
# $bcc = 'kent.mamalias@metrogaisano.com';
		
$subject = 'In Stock Report';
$msgbody_file = 'message.txt';

$attachment_file = "INSTOCK_REPORT.xlsx";

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
Content-Type: text/plain; charset="iso-8859-1"
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

sub mail3 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = ' carmelita.intia@metrogaisano.com, consorcia.mullon@metrogaisano.com, contessa.fernandez@metrogaisano.com, delia.jakosalem@metrogaisano.com, editha.cabriles@metrogaisano.com, edna.prieto@metrogaisano.com, jecil.cumayas@metrogaisano.com, jennifer.yu@metrogaisano.com, lorena.madraga@metrogaisano.com, maryann.delarama@metrogaisano.com, maryjoy.montes@metrogaisano.com, mecelle.quimbo@metrogaisano.com, mirasol.barcoma@metrogaisano.com, nenita.cabigon@metrogaisano.com, teresita.manatad@metrogaisano.com, vilma.paner@metrogaisano.com, melinda.uy@metrogaisano.com, jgeniston@metro.com.ph, cartes@metro.com.ph, hcaberte@metro.com.ph, mvilla@metro.com.ph, mcabungcal@metro.com.ph, mcombinido@metro.com.ph, april.agapito@metrogaisano.com, jordan.mok@metrogaisano.com, marlita.portes@metrogaisano.com, advento.resma@metrogaisano.com, arlene.te@metrogaisano.com, armando.pitogo@metrogaisano.com, christine.lanohan@metrogaisano.com, delvie.pitogo@metrogaisano.com, emillie.ponsica@metrogaisano.com ';
$cc = 'cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, jerson.roma@metrogaisano.com, bermon.alcantara@metrogaisano.com, eljie.laquinon@metrogaisano.com, nilynn.yosores@metrogaisano.com, ryanneil.dupay@metrogaisano.com, anafatima.mancho@metrogaisano.com';
$bcc = 'kent.mamalias@metrogaisano.com';
		
# $to = 'kent.mamalias@metrogaisano.com';
# $bcc = 'kent.mamalias@metrogaisano.com';
		
$subject = 'In Stock Report';
$msgbody_file = 'message.txt';

$attachment_file = "INSTOCK_REPORT.xlsx";

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
Content-Type: text/plain; charset="iso-8859-1"
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

sub mail4 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = ' jennifer.moreno@metrogaisano.com, jessica.gaisano@metrogaisano.com, keith.poblete@metrogaisano.com, micah.alvarado@metrogaisano.com, michelle.someros@metrogaisano.com, peachy.aquino@metrogaisano.com, ryan.uson@metrogaisano.com, sheen.ducay@metrogaisano.com, annie.desuyo@metrogaisano.com, marlit.ignacio@metrogaisano.com, mildred.quinones@metrogaisano.com, rosemarie.saravia@metrogaisano.com, rowena.conde@metrogaisano.com, tessie.baldezamo@metrogaisano.com, alain.reyes@metrogaisano.com, diana.almagro@metrogaisano.com, angelito.dublin@metrogaisano.com, arlene.yanson@metrogaisano.com, charm.buenaventura@metrogaisano.com, teena.velasco@metrogaisano.com, cristy.sy@metrogaisano.com, fidela.villamor@metrogaisano.com, glenda.navares@metrogaisano.com, jonalyn.diaz@metrogaisano.com, josemarie.graciadas@metrogaisano.com, judy.gilo@metrogaisano.com, cristina.de_asis@metrogaisano.com, mariajoana.cruz@metrogaisano.com, ricky.tulda@metrogaisano.com, roy.igot@metrogaisano.com ';
$cc = 'cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, jerson.roma@metrogaisano.com, bermon.alcantara@metrogaisano.com, eljie.laquinon@metrogaisano.com, nilynn.yosores@metrogaisano.com, ryanneil.dupay@metrogaisano.com, anafatima.mancho@metrogaisano.com';
$bcc = 'kent.mamalias@metrogaisano.com';
		
# $to = 'kent.mamalias@metrogaisano.com';
# $bcc = 'kent.mamalias@metrogaisano.com';
		
$subject = 'In Stock Report';
$msgbody_file = 'message.txt';

$attachment_file = "INSTOCK_REPORT.xlsx";

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
Content-Type: text/plain; charset="iso-8859-1"
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

sub mail5 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = ' may.sasedor@metrogaisano.com, donna.fernando@metrogaisano.com, victoria.ferolino@metrogaisano.com, wendel.gallo@metrogaisano.com, ronald.dizon@metrogaisano.com, alexander.tejedor@metrogaisano.com ';
$cc = 'cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, jerson.roma@metrogaisano.com, bermon.alcantara@metrogaisano.com, eljie.laquinon@metrogaisano.com, nilynn.yosores@metrogaisano.com, ryanneil.dupay@metrogaisano.com, anafatima.mancho@metrogaisano.com';
$bcc = 'kent.mamalias@metrogaisano.com';
		
# $to = 'kent.mamalias@metrogaisano.com';
# $bcc = 'kent.mamalias@metrogaisano.com';
		
$subject = 'In Stock Report';
$msgbody_file = 'message.txt';

$attachment_file = "INSTOCK_REPORT.xlsx";

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
Content-Type: text/plain; charset="iso-8859-1"
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

sub read_file {

my( $filename, $binmode ) = @_;
my $fh = new IO::File;
$fh->open("<".$filename) or die "Error opening $filename for reading - $!\n";
$fh->binmode if $binmode;
local $/;
<$fh>
	
}








