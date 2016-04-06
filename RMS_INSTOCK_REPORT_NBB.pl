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

my $workbook = Excel::Writer::XLSX->new("INSTOCK_REPORT_NBB.xlsx");
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

printf "IN STOCK REPORT, NBB \n";

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
&h9;
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

# &new_sheet($sheet = "ByDept-Whse", $comment = "WAREHOUSE");
# &w80001;
# &w80011;
# &w80031;
# &w80041;
# &w80051;
# &w80061;
# &calc_tot_region($total_label = 'TOTAL WHSE', $end = 0);

$workbook->close();
$dbh_csv->disconnect;

&mail1;	
&mail2;	
&mail3;	
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
sub h9 { $store = "H9"; $loc = "'6009'"; &call_div;	}
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

# &div1;
# &div2;
# &div3;
# &div4;
# &div5;
# &div6;
# &calc($code = "Department Store", $ds = $a); # Department Store Total
&div8;
&div9;
# &div10;
# &div11;
&calc($code = "Supermarket"); # Supermarket Total
#&calc7; #TOTAL

$test = 0, $a = 0, $ds = 0, $e = 0, $counter=0; #RE INITIALIZE VARIABLES
$col += 3;

}

sub div1 {

$div_name = "Appliance and Home Improvement";
$g1 = '2530'; $g2 = '0000'; $g3 = '0000'; $g4 = '0000'; $g5 = '0000'; $g6 = '0000'; $g7 = '0000'; $g8 = '0000'; $g9 = '0000'; $g10 = '0000'; 
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
$g1 = '3520'; $g2 = '3530'; $g3 = '3510'; $g4 = '5010'; $g5 = '7010'; $g6 = '7020'; $g7 = '3560'; $g8 = '3570'; $g9 = '3580'; $g10 = '0000'; 
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

$worksheet->write(0, 3, "TOTAL " . $comment , $bold);
$worksheet->write(1, 3, "As of " . $day . '-' . $month_to_text . '-' .$year, $bold2);
$worksheet->merge_range( 4, 3, 5, 3, 'DIV', $subhead );
$worksheet->merge_range( 4, 5, 5, 5, 'CODE', $subhead );
$worksheet->merge_range( 4, 6, 5, 6, 'DEPARTMENT', $subhead );

}

sub query_markdown {

$table = 'instock_nbb.csv';

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

# if($end eq 1){
	foreach my $c( 6..$count ){
		my $sumCount = '=SUMIFS('.xl_rowcol_to_cell($c,7).':'.xl_rowcol_to_cell($c,$col-2).','. xl_rowcol_to_cell(5,7).':'. xl_rowcol_to_cell(5,$col-2).',"COUNT")';
			$worksheet->write( $c, $col, $sumCount, $headNumber );
		my $sumStock = '=SUMIFS('.xl_rowcol_to_cell($c,7).':'.xl_rowcol_to_cell($c,$col-2).','. xl_rowcol_to_cell(5,7).':'. xl_rowcol_to_cell(5,$col-2).',"IN STOCK")';
			$worksheet->write( $c, $col+1, $sumStock, $headNumber );		
		my $pct = '=IFERROR('. xl_rowcol_to_cell( $c, $col+1 ). '/' . xl_rowcol_to_cell( $c, $col ).', )' ;
			$worksheet->write( $c, $col+2, $pct, $headN );
	}
# }
# elsif($end eq 0){
	# foreach my $c( 6..$count+2 ){
		# my $sumCount = '=SUMIFS('.xl_rowcol_to_cell($c,$vis).':'.xl_rowcol_to_cell($c,$col-1).','. xl_rowcol_to_cell(5,$vis).':'. xl_rowcol_to_cell(5,$col-1).',"COUNT")';
			# $worksheet->write( $c, $col, $sumCount, $headNumber );
		# my $sumStock = '=SUMIFS('.xl_rowcol_to_cell($c,$vis).':'.xl_rowcol_to_cell($c,$col-1).','. xl_rowcol_to_cell(5,$vis).':'. xl_rowcol_to_cell(5,$col-1).',"IN STOCK")';
			# $worksheet->write( $c, $col+1, $sumStock, $headNumber );		
		# my $pct = '=IFERROR('. xl_rowcol_to_cell( $c, $col+1 ). '/' . xl_rowcol_to_cell( $c, $col ).', )' ;
			# $worksheet->write( $c, $col+2, $pct, $headN );
	# }
# }

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
 open my $fh, ">", "instock_nbb.csv" or die "instock_nbb.csv: $!";

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
	  STOCK_ON_HAND, DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG
	FROM REPL_ITEM_LOC REPL
		LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
		LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
		LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
		LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
		LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
		LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
	WHERE LOC.STATUS IN ('A')  AND 
	(
	(REPL.LOCATION IN (2004,2005,2006,2007,2009,2012,2013,3007,4004,6002) AND REPL.ITEM IN ( 10331962,10302110,10516960,9083975,10071291,6393335,2068268,2089607,2052052,3132173,2071664,9879223,10627605,2068336,2068305,3758632,6393366,3132166,2051963,2083223,6171261,10739487,10517402,9603008,10573829,10343097,9288596,9378636,2012308,9604234,2099224,10510572,2083247,9378666,9603005,3762387,2090092,10677353,10331991,9083978,9834562,10091953,2089966,1999785,9617526,2042831,3118603,2063737,5544264,7173233,2094106,2046563,3118610,3118634,8404954,9287582,2104485,9985068,9058949,9457103,5244904,6256548,15003212,2088853,9420841,3786420,9104215,2112411,9192574,2112459,2063744,2065090,9279558,9056027,4082392,9985071,2077017,2085104,2088785,9245716,2089126,1992885,9436232,2112428,10666074,2016269,9790857,2088860,2088792,6447779,15003036,8404794,2088914,9616898,9219375,10762541,6256500,5672202,10762542,9013487,2043739,8202529,5672189,9419780,8713155,2094359,2094342,10137848,9013488,10175391,8502599,3313855,8202512,9858798,2035000,8956972,2060347,8956989,1991680,8432315,2003429,2119489,2003603,2035604,2059099,2003436,2035314,8369840,2003528,9191437,1991758,10044690,9257330,9413022,10075820,10154458,10154459,9798593,9128159,9798596,9304829,3794753,9628055,2069180,9945757,9413018,9945758,10730251,8779199,2069197,9938810,3322376,2040219,9229174,2097497,10752794,2108650,9592783,9628042,2099217,10640020,2097473,9592734,7081873,9772645,2099392,2070209,9229170,9516657,10225270,2089867,2099064,10082124,5419357,5385980,8201966,10082128,10673057,3356647,9354924,3356630,9111641,9111655,2106441,10750736,9612599,9203918,10269411,2101378,10228601,2002941,7863875,2098999,5239047,6007867,9523786,15000321,15002228,15002229,9872650,9872651,2095929,2115597,2116273,2450070,2091631,1840223,9243002,9261923,10089285,2087047,9499921,10040740,9985192,2091846,9592985,2091938,2091839,9597891,9788582,9788570,10165161,9716380,15003012,15003011,2091624,9462912,9897461,8790606,9057602,9034647,10430333,15001904,2004310,2004327,9780086,9893598,9893594,9617966,9850872,10022702,9292448,9363194,2096407,2096667,10737453,9507012,9516424,9015703,10430315,10430333,10640772,10640771,10195624,9495928,10334147,10640769,9631095,10244761,9894887,9837032,9842352,9341571,9755384,9920013,10215077,9890643,9920008,9183052,9380279,9534877,9919994,9183065,1999778,10017759,10017901,10774384,10701126,10020873,9963896,10343284,10629075,9887384,9917884,9377775,10122103,10629076,10075303,10545936,10085788,10508099,10612338,10178007,10505504,10485489,10021145,9377770,9308960,10507565,10031518,10508098,9904404,10734466,9787751,10629077,10178012,9548130,10007259,10148959,10456639,9787764,10629079,10343298,10179040,2086453,10141763,9660462,9787779,9788271,10133888,2086460,2086446,2086439,10252423,10235136,10734467,10715163,9215824,10271131,9268660,9244421,10629083,2067544,10285308,10285307,9853613,9425602,10285304,9853597,9853623,10732327,3365281,9014625,10732321,10285266 )) OR 
	(REPL.LOCATION IN (2003,2008,2010,2011,3001,3002,3003,3004,3005,3006,3009,4003,6001,6003,6005,6010) AND REPL.ITEM IN (10516960,9083975,10627605,10071291,10627604,10750783,9604234,6768300,10302110,10517386,6393335,9879223,9083978,10340354,2089607,2071664,2099224,10517397,10573833,2047980,10573831,7928482,10331962,2068268,2089966,2014173,10205505,1999785,2099767,10517402,10517399,10574018,7173233,9118909,15003212,5544264,2096742,3118603,5272112,15003579,2089126,3781357,9457103,9509344,3118610,1989328,9279558,2088785,10143867,5244904,2077017,9025360,9058943,7385018,9637407,2046563,6049010,7385230,8404954,2063737,10143731,2067575,2067582,2092928,2044323,2092904,9142746,2099897,2035000,2119489,2116525,2092935,2035604,2092911,2044316,5402625,2035314,9236135,9257330,8956972,2116518,2036311,3772997,9938810,10752795,10752794,3322376,10135507,2040219,10264849,2089386,2099217,10751569,9229174,9772645,10751568,9656160,9285894,10170210,9909594,3203873,9583918,10082124,2112701,9220361,7457548,5385980,5419814,2112640,10738766,2095424,8201966,5419357,5419890,10719005,10719000,10082128,2112657,9413022,10506394,9413021,3794753,10151719,15001916,10716847,9719678,8779199,10151714,10730251,9192668,2069180,15001917,9285314,10144094,10699762,9099661,10228621,9059889,10119794,9598807,9727183,9099310,9732066,8750341,10217611,10657317,10673057,9354924,3356654,3356647,10623494,9354914,3356630,3311189,3311769,3311790,2088792,2088860,8404794,6447779,15003036,9070024,9070021,2088914,9070026,9216858,9616898,2068510,2099392,2070209,9229170,10225270,2070216,2089867,7215469,2099125,2091693,2099064,3320747,3320730,9166462,9101411,3316016,10269411,3344729,9951697,3318553,9111655,9111641,9020346,2101699,2101880,101414102,2002064,5239047,9523786,15002228,2095929,9058920,9552315,10117970,9116097,9780086,9678254,2091631,1840223,9858329,9869425,9470326,10066546,10617133,9797852,9631095,9499921,9853692,9985192,2091846,9592985,2091839,9788582,9788570,10165161,9716380,15003011,10001076,9897462,9897461,9057602,9034647,9981874,7759376,2004310,2004334,2004327,10637570,9034649,10022702,9292447,9160194,9292448,9363194,10216301,2096407,9894887,2096667,10737453,9507016,9507012,10430481,10430482,10430315,10430302,10430333,10185632,10195624,2109657,10334147,10385290,9631093,9689350,10179705,9894887,9837032,9111616,9755407,7592829,9058929,7592805,10165797,8418579,9497411,10165789,9100338,9058931,15003295,10002253,9058928,15003384,9871562,9112521,15002863,8904829,9948857,15002862,9919994,8904843,7592843,10263564,9948856,9530162,15000380,9111608,10017759,10017901,10020873,10629075,10178007,2085845,10774384,2085838,2091914,10701126,10242216,2086446,10178012,10343284,10690146,10242203,9963896,9787764,10141763,9787751,9788271,2086439,2091907,2067544,10650911,2086453,2085821,10133894,10629076,10178035,10456639,9788272,2071909,9425602,9163010,2103945,9853597,9853613,9853623,2120003,2086323,9086720,9117619,2103952,10116387,3365281,10285302,2104072,3359662,8899859,7503085,1985375)) OR
	(REPL.LOCATION IN (2001,2002,4002,6004) AND REPL.ITEM IN (10516960,10627605,10302110,10071291,6768300,9604234,10750783,10517386,9083975,10627604,10517397,2089607,10573831,10331962,10517402,2035550,10573833,6393335,9879223,2099224,2042497,10574018,2014173,10517399,2068268,9940408,9612895,9378636,9083978,2099767,1999785,9083977,7928482,10340354,2071664,2089966,10739486,2060958,9378666,9013457,2008837,10205505,10702195,9165981,10573832,9506226,10205498,10667465,6393366,10739487,8840080,2060958,10673808,101412767,2112701,9220361,5385980,7457548,5419814,10082124,10738766,2112640,5419890,2095424,10738768,2112657,2112718,5419357,5419166,10226324,10043024,9337300,10082128,8608055,10719005,10719000,7857461,6872212,9337282,5419616,10748879,9750998,8201966,8609229,5213689,10691912,9320968,10147127,10719007,7884207,8921253,8921260,10509465,10719002,5419746,7884221,3366356,2119090,9581242,5875771,5764563,10113551,3758472,9118909,7173233,2096742,15003212,3781357,5272112,15003579,5544264,3118603,2089126,9509344,7385018,9457103,9279558,9637407,7385230,2025575,2064932,9058943,2088785,1989328,2094106,10143731,2061702,2003320,2032160,7530982,2046563,5263073,9118033,8404954,2077017,2067575,2092904,2092928,2067582,101505904,9142746,6023751,8956972,2035000,2059976,15002273,2092935,8956989,2035314,9236135,2099897,15002274,2978185,6686642,2116525,2092911,9142740,2067759,2116518,9118030,2003436,2099859,2119489,2035604,10044690,2003429,9938810,10752794,10135507,10752795,2040219,10264849,3322376,9229174,2089386,9772645,2089416,9656160,10170210,10751569,9612265,2099217,2058337,9909594,2090115,7273445,10751568,2089409,9857410,10755863,9583982,2108650,2097473,8404794,2088792,2088860,15003036,6447779,9219375,9616898,2068510,9070021,9070026,9070024,2088914,101478858,2043739,101414102,2068626,2002064,3334775,3814581,3334850,9413022,10506394,5845361,3794753,8779199,9413021,2104768,9438033,9128164,9192668,9304829,2104652,2069180,9413018,9172002,15003006,2065502,9531053,9285314,2094670,8670106,10699762,9727183,10144094,2055701,9052732,8669797,9059889,2055718,10181461,9099661,8750341,10657317,2077239,8765901,2070209,2099392,9229170,2089867,10225270,2091693,2070216,2099125,10673057,10623494,9808939,9354924,3356647,9354914,3356654,3316016,9166462,3344729,9951697,9893963,10269411,9111655,9111641,10163810,2101699,10217611,2099194,9733454,2022499,2024219,15002228,15002229,2095936,2095929,9552315,10117970,2115597,9678254,9733450,2091631,10692343,1840223,9261923,10031646,10212653,10031647,10031650,10190102,10743021,9499921,2091624,2092294,10185561,9985192,2091846,9592985,2091938,2091839,9995112,9995108,9995110,9788582,9788570,9716380,15003011,8416117,10001076,9462912,9897462,9897462,9897461,9057602,9057602,9034647,9981874,2096391,7759376,2004310,2004334,2004327,9897463,101498826,1993721,1986297,9780086,10022702,9292447,9292448,9363194,9363187,9362741,9766479,9766555,10216301,2096407,8202055,10737460,2096667,10737453,9507012,10430481,10430480,10430482,10430315,10430302,10430333,10185632,10745576,10195624,9244183,10428799,10428801,10428802,10334147,10385248,10385290,10195646,10564581,9631093,9631095,9689350,10179709,10179705,10244766,9894887,9837032,10346741,7592829,7592805,15003295,9111616,15003384,9058929,8904843,7592775,8904829,10165789,7592751,9755384,9183620,15002863,9100338,9530162,10165797,9183622,9497411,15000380,7592843,9058928,15002862,9425613,10485506,9425620,9425612,10485505,9058931,2031590,15003385,9755407,9111608,7592782,8904812,9247990,8418579,10017759,10017901,10020873,10178007,10774384,10629075,2067544,10178012,2085845,10242216,2086446,10343284,2086453,9787751,9788271,2085838,9963896,10650911,10141763,10242203,10690146,9215824,10456639,2091914,2086460,2086439,10701126,10629076,2091907,9788272,10188432,7116650,9787764,9787779,2067551,9787778,10178035,9377775,9887384,2114262,5232598,9777626,9268660,9253440,10485489,10715163,10343298,10021145,15000362,10179040,9777631,10650913,9258011,10545936,2113968,10343299,10242229,10629077,10738802,10164363,2071909,9425602,9163010,10285302,2120003,3359662,10285266,10508100,2086323,10285308,9853613,9853623,10616659,9853597,10122798,3365281,10116387,10285307,10116404,2091464,2058399,2103945,1985375))
	)
	)TBL 
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
 close $fh or die "instock_nbb.csv: $!";
 
$dbh->disconnect;

}

sub mail1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = ' cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, victoria.abasolo@metrogaisano.com, bermon.alcantara@metrogaisano.com, eljie.laquinon@metrogaisano.com, nilynn.yosores@metrogaisano.com, ryanneil.dupay@metrogaisano.com, anafatima.mancho@metrogaisano.com, emily.silverio@metrogaisano.com, luz.bitang@metrogaisano.com ';
$bcc = 'kent.mamalias@metrogaisano.com, rex.cabanilla@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com';

# $to = ' kent.mamalias@metrogaisano.com';
		
$subject = "In Stock Report, NBB";
$msgbody_file = 'message.txt';

$attachment_file = "INSTOCK_REPORT_NBB.xlsx";

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

$to = ' ace.olalia@metrogaisano.com, alain.reyes@metrogaisano.com, alma.espino@metrogaisano.com, angeli_christi.ladot@metrogaisano.com, angelito.dublin@metrogaisano.com, arlene.yanson@metrogaisano.com, augosto.daria@metrogaisano.com, charm.buenaventura@metrogaisano.com, teena.velasco@metrogaisano.com, cristy.sy@metrogaisano.com, diana.almagro@metrogaisano.com, dinah.ramirez@metrogaisano.com, edgardo.lim@metrogaisano.com, edris.tarrobal@metrogaisano.com, fidela.villamor@metrogaisano.com, genaro.felisilda@metrogaisano.com, genevive.quinones@metrogaisano.com, glenda.navares@metrogaisano.com, jacqueline.cano@metrogaisano.com, joefrey.camu@metrogaisano.com, jonalyn.diaz@metrogaisano.com ';
$bcc = 'kent.mamalias@metrogaisano.com';

# $to = ' kent.mamalias@metrogaisano.com';
		
$subject = "In Stock Report, NBB";
$msgbody_file = 'message.txt';

$attachment_file = "INSTOCK_REPORT_NBB.xlsx";

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

$to = ' josemarie.graciadas@metrogaisano.com, jovany.polancos@metrogaisano.com, judy.gilo@metrogaisano.com, julie.montano@metrogaisano.com, kathlene.procianos@metrogaisano.com, limuel.ulanday@metrogaisano.com, cristina.de_asis@metrogaisano.com, mariajoana.cruz@metrogaisano.com, may.sasedor@metrogaisano.com, michelle.calsada@metrogaisano.com, policarpo.mission@metrogaisano.com, rex.refuerzo@metrogaisano.com, ricky.tulda@metrogaisano.com, ronald.dizon@metrogaisano.com, roselle.agbayani@metrogaisano.com, rowena.tangoan@metrogaisano.com, roy.igot@metrogaisano.com, tessie.cabanero@metrogaisano.com, victoria.ferolino@metrogaisano.com, wendel.gallo@metrogaisano.com, juanjose.sibal@metrogaisano.com, julie.montano@metrogaisano.com, noli.lee@metrogaisano.com, vivian.ablang@metrogaisano.com, roselle.agbayani@metrogaisano.com, irene.montemayor@metrogaisano.com ';
$bcc = 'kent.mamalias@metrogaisano.com';

# $to = ' kent.mamalias@metrogaisano.com';
		
$subject = "In Stock Report, NBB";
$msgbody_file = 'message.txt';

$attachment_file = "INSTOCK_REPORT_NBB.xlsx";

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








