START:

use DBI;
use DBD::Oracle qw(:ora_types);
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Utility;
use Text::CSV_XS;
#use DateKey_ARC;
#use DBConnector;
use Win32::Job;
use Getopt::Long;
use IO::File;
use MIME::QuotedPrint;
use MIME::Base64;
use Mail::Sendmail;
use Date::Calc qw( Today Add_Delta_Days Month_to_Text);

($year,$month,$day) = Today();
$month_to_text = Month_to_Text($month);

##on apr 7, jerson provided a new list of items

# $test_query = qq{ SELECT CASE WHEN EXISTS (SELECT SEQ_NO, ETL_SUMMARY, VALUE, ARC_DATE FROM ADMIN_ETL_SUMMARY WHERE TO_DATE(ARC_DATE, 'DD-MON-YY') = TO_DATE(SYSDATE,'DD-MON-YY')) THEN 1 ELSE 0 END STATUS FROM DUAL };

# $tst_query = $dbh->prepare($test_query);
# $tst_query->execute();

# while ( my $x =  $tst_query->fetchrow_hashref()){
	# $test = $x->{STATUS};
# } 
$test = 1;
if ($test eq 1){

	# $date = qq{ 
	# SELECT WEEK_NUMBER_THIS_YEAR, DATE_KEY, TO_CHAR(DATE_FLD, 'DD Mon YYYY') DATE_FLD, WEEK_ST_DATE_KEY, WEEK_END_DATE_KEY 
	# FROM DIM_DATE 
	# WHERE DATE_FLD = (SELECT TO_DATE(VALUE,'YYYY-MM-DD') FROM ADMIN_ETL_SUMMARY)
	 # };

	# my $sth_date_1 = $dbh->prepare ($date);
	 # $sth_date_1->execute;

	# while (my $x = $sth_date_1->fetchrow_hashref()) {
		# $as_of = $x->{DATE_FLD};
	# }
	
	$workbook = Excel::Writer::XLSX->new("INSTOCK_REPORT_INTERNATIONAL_v2_US.xlsx");
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
	$headN = $workbook->add_format( border => 1, align => 'center', valign => 'center', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
	$headD = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 9, bg_color => $abo, bold => 1 );
	$headPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
	$headNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 3, bg_color => $abo, bold => 1 );
	$headNumber = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 3, bg_color => $abo, bold => 1 );
	$head = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 11, num_format => 9, bg_color => $abo, bold => 1 );
	$subhead = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, bg_color => $ponkan, bold => 1 );
	$bodyN = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
	$bodyPct = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
	$bodyNum = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 3,  bold => 1);
	$body = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, bg_color => $sky, num_format => 9,  bold => 1);
	$subt = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 9);
	$down = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', size => 10, num_format => 9, bg_color => $pula );

	# printf "Test ETL Status = ". $test ." \nArc BI Sales Performance Part 1\nGenerating Data from Source \n";
	
	# $tst_query->finish();
	
	&generate_csv;
	
	&new_sheet($sheet = "Store");
	
	&call_summary_division($area_flg1 = 3, $area_flg2 = 3);
	
	&new_sheet($sheet = "Warehouse");
	
	&call_summary_division($area_flg1 = 1, $area_flg2 = 2);
	
	$workbook->close();
	$dbh_csv->disconnect;
	
	&mail1;
	&mail2;
	&mail3;
	
	exit;
	
}

else{
	print "Test Status = ". $test ." \nETL still running\n";
	
	$tst_query->finish();
	
	sleep(600);
	
	goto START;
}
 
 
sub call_summary_division {

##=============China
$a = 0; $vis = 7; $col = 7; $test = 1; $stopper = 0; 

&heading;

$country = 'China_HK_Taiwan', $region = 'China/HK/Taiwan', $a += 6, $counter = 0, $len = 0, $loc_pt = $a-2;

$worksheet->set_column( $col, $col+1, 7 );
$worksheet->set_column( $col+2, $col+2, 8 );

foreach my $i ("COUNT", "IN STOCK", "%") {
	$worksheet->write($a-1, $col++, $i, $subhead);
}

$col -= 3;

&query_summary_division;

$test = 0, $a = 0, $counter = 0, $stopper = 1, $col += 3;



$worksheet->set_column( 7, 8, undef, undef, 1 );
$worksheet->set_column( 10, 11, undef, undef, 1 );
$worksheet->set_column( 13, 14, undef, undef, 1 );
$worksheet->set_column( 16, 17, undef, undef, 1 );
$worksheet->set_column( 19, 20, undef, undef, 1 );
$worksheet->set_column( 22, 23, undef, undef, 1 );
$worksheet->set_column( 25, 26, undef, undef, 1 );
$worksheet->set_column( 28, 29, undef, undef, 1 );
$worksheet->set_column( 31, 32, undef, undef, 1 );

}



sub new_sheet{

$worksheet = $workbook->add_worksheet($sheet);
$worksheet->set_zoom(85);
$worksheet->set_paper( 8 );
$worksheet->center_horizontally();
$worksheet->set_print_scale( 100 );
$worksheet->set_margins( 0.05 );
$worksheet->conditional_formatting( 'F9:V2000',  { type => 'cell', criteria => '<', value => 0, format => $down } );
$worksheet->set_column( 1, 3, undef, undef, 1 );
# $worksheet->set_column( 1, 2, 3 );
# $worksheet->set_column( 3, 4, 4 );
$worksheet->set_column( 5, 5, undef, undef, 1 );
$worksheet->set_column( 6, 6, 21 );

}

sub heading {

$worksheet->write(0, 4, "Replenishment In-Stock " . $comment , $bold);
$worksheet->write(1, 4, $day . '-' . $month_to_text . '-' .$year, $bold2);

$worksheet->merge_range( 4, 2, 5, 2, 'Type', $subhead );
$worksheet->merge_range( 4, 3, 5, 3, 'Type', $subhead );
# $worksheet->merge_range( 4, 4, 5, 4, 'Type', $subhead );
# $worksheet->merge_range( 4, 5, 5, 5, 'Code', $subhead );
$worksheet->merge_range( 4, 4, 5, 6, 'Location', $subhead );

}
 

sub query_summary_division {

$table = 'instock_us.csv';

$dbh_csv = DBI->connect("dbi:CSV:f_dir=$ENV{HOME}/csvdb;f_ext=.csv;f_encoding=utf8;csv_eol=\n;csv_sep_char=",";csv_quote_char=\";csv_escape_char=\\;csv_class=Text::CSV_XS;csv_null=1") 
			or die $DBI::errstr;
			
$sls1 = $dbh_csv->prepare (qq{SELECT country, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
									 FROM $table 
									 WHERE country = '$country' and (area <> '$area_flg1' and area <> '$area_flg2')
								GROUP BY country
							 }); $sls1->execute();

while(my $s = $sls1->fetchrow_hashref()){
	$country = $s->{country};
	
	$sls2 = $dbh_csv->prepare (qq{SELECT area, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
									 FROM $table 
									 WHERE country = '$country' and (area <> '$area_flg1' and area <> '$area_flg2')
									 GROUP BY area
									 ORDER BY area
								}); $sls2->execute();
	
	$mgc_counter = $a;
	while(my $s = $sls2->fetchrow_hashref()){
		$area = $s->{area};
				
		$sls3 = $dbh_csv->prepare (qq{SELECT store, store_name, SUM(tot_repl_items) AS tot_items, SUM(repl_with_soh) AS rep_items
									 FROM $table 
									 WHERE country = '$country' AND area = '$area'
									 GROUP BY store, store_name
									 ORDER BY store
									}); $sls3->execute();
									
		while(my $s = $sls3->fetchrow_hashref()){
		
			if($stopper eq 0){ $worksheet->merge_range( $a, 4, $a, 6, $s->{store}.'-'.$s->{store_name}, $desc );}
			
			$worksheet->write($a,$col, $s->{tot_items},$border1); 
			$worksheet->write($a,$col+1, $s->{rep_items},$border1);
				if ($s->{tot_items} <= 0){
					$worksheet->write($a,$col+2, "",$subt); 				}
				else{
					$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$subt); 				}
			
			$counter = 0; #RESET dept_counter	
			$a++; #INCREMENT VARIABLE a
		}

		$worksheet->write($a,$col, $s->{tot_items},$bodyNum); 
		$worksheet->write($a,$col+1, $s->{rep_items},$bodyNum);
		if ($s->{tot_items} <= 0){
			$worksheet->write($a,$col+2, "",$bodyPct); 		}
		else{
			$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$bodyPct); 		}
		
		if($stopper eq 0){
			$worksheet->merge_range( $a, 4, $a, 6, 'Subtotal', $bodyN );}
		
		$a++; #INCREMENT VARIABLE a
	}
	
	$worksheet->write($a,$col, $s->{tot_items},$headNumber); 
	$worksheet->write($a,$col+1, $s->{rep_items},$headNumber);
	if ($s->{tot_items} <= 0){
		$worksheet->write($a,$col+2, "",$headPct); 	}
	else{
		$worksheet->write($a,$col+2, '=IF(ISERROR('. $s->{rep_items}/$s->{tot_items} .'),"",('.$s->{rep_items}/$s->{tot_items}. '))',$headPct); 	}

}

if($stopper eq 0){		 	$worksheet->merge_range( $a, 2, $a, 6, 'TOTAL', $headN ); }

$worksheet->merge_range( $loc_pt, $col, $loc_pt, $col+2, $region, $subhead );

$sls1->finish();
$sls2->finish();
$sls3->finish();
#$sls4->finish();

$len = $a, $counter = 0;

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
 open my $fh, ">", "instock_us.csv" or die "instock_us.csv: $!";

$test = qq{ 
SELECT SGD.GROUP_NO, SGD.STORE, SGD.STORE_NAME, CASE WHEN SGD.STORE IN ('2004','2005','2006','2007','2009','2012','2013','2223','3007','3012','4004','6002','6012','6013') THEN '2'     WHEN SGD.STORE IN ('80001','80011','80031','80041','80051','80061') THEN '3' ELSE '1' END AS AREA, STK.TOT_REPL_ITEMS, STK.REPL_WITH_SOH, 'US' AS COUNTRY FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME 
			FROM
				(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (3010,6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
				(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
					FROM DEPS D 
					  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
					  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
					  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
			WHERE S.STORE IN ('2001','2002','2003','2008','2010','2011','4002','4003','3009','6001','6003','6005','6010','3001','3002','3003','3004','3005','3006','6004','6009','2004','2005','2006','2007','2009','2012','2013','2223','3007 ','3012','4004','6002','6012','80001','80011','80031','80041','80051','80061')
			GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME
		)SGD
	LEFT JOIN
	(SELECT GROUP_NO, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
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
			WHERE --LOC.STATUS IN ('A') AND  REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE)
				REPL.ITEM IN ('10010486','10011585','10011588','10016996','10017016','10017017','10031434','10050111','10050385','10050749','10051870','10053030','10054794','10055296','10055334','10055343','10055528','10055688','10055751','10055848','10056144','10056213','10056229','10056239','10069384','10072990','10075340','10075343','10082667','10084103','10084107','10095779','10096601','10099488','10099626','10102799','10104417','10112401','10117631','10117638','10129482','10133140','10133482','10148495','101499205','101501022','101564159','10164932','10164938','10164940','10164969','10164974','10166912','10166915','10166919','10166973','10171677','10171680','10171726','10171737','10171753','10172172','10172180','10172182','10176926','10176932','10176934','10176939','10177915','10177968','10177979','101807840','101807942','101809876','101810273','101810338','101820420','101820434','101820460','101820493','101820953','10186356','101869045','101905135','101905304','10194409','10194412','101973814','10205002','10205292','10205766','10220212','10220214','10220215','10220216','10220217','10246922','10300880','10300900','10302012','10325076','10328713','10328715','10328718','10342846','10353028','10362081','10362083','10362085','10400925','10432458','10432643','10450162','10521826','10521827','10521854','10521857','10521863','10521915','10521925','10521928','10521934','10521939','10522014','10522016','10522019','10522021','10522023','10522025','10555786','10555787','10555799','10555830','10556041','10559430','10564634','10564641','10564662','10564676','10617981','10617982','10617983','10618233','10618253','10618262','10618270','10618272','10618719','10618902','10618922','10618998','10619031','10619032','10619033','10619034','10619042','10619043','10619044','10628577','10628591','10628620','10628651','10628655','10628669','10628683','10628688','10628704','10628706','10637065','10637066','10637067','10659618','10659619','10659620','10659621','10659622','10659623','10659625','10659626','10659627','10659628','10659630','10659635','10660040','10660042','10661550','10663157','10663158','10663167','10669475','10669485','10669492','10669511','10723498','10752126','15001554','15001564','15001586','15003172','15003961','15003971','1986358','1992038','1997668','2062334','2083803','3753002','3756362','3756836','5228508','5228836','5228881','5228966','5230044','5230051','5230341','5230600','5241613','5244089','5244997','5245000','5245055','5263851','5280254','5294923','5296729','5484041','5488544','5488582','5489572','5489589','5489596','5489732','5490998','5492633','5492671','5492701','5492763','5492831','5493005','5493012','5493050','5497751','5501663','5503063','5542437','5542444','5597123','5597130','5645206','5645213','5645800','5646050','5646067','5646081','5653744','5653775','5655182','5655663','5656554','5656561','5658053','5658077','5658411','5658725','5684571','5684625','5685165','5688401','5688913','5691210','5691319','5691449','5693375','5693382','5693573','5721009','5726646','5726677','5727575','5728039','5728084','5731602','5731619','5793525','5803934','5909025','5912254','5989713','5996537','5996568','5996650','5996858','5997480','6010126','6010508','6010584','6011383','6012175','6022181','6022457','6106782','6110857','6111045','6111717','6416713','6505592','6535940','6540951','6541002','6541279','6541460','6543426','6638696','6840914','6840952','6841782','6844394','6846886','6846916','6847029','6888152','6888237','6888480','6905392','6905552','7164231','7167799','7167935','7168222','7168246','7168260','7168710','7182945','7183294','7280023','7283383','7316531','7317996','7318627','7319815','7326219','7327933','7338779','7341892','7375842','7379253','7382765','7385834','7444654','7502422','7683299','7741555','7741999','7743863','7786570','7864315','8118516','8118578','8156556','8156587','8157867','8266101','8273321','8274250','8278654','8283054','8409270','8436351','8437778','8438997','8564542','8564580','8847621','8850423','9026044','9026055','9026800','9026829','9026962','9026964','9027068','9027072','9027446','9040887','9040890','9043945','9074467','9075436','9076398','9076415','9095165','9095183','9095227','9095339','9095557','9095687','9113977','9123959','9124117','9124160','9187417','9187432','9187437','9187446','9187467','9187472','9187481','9188662','9188786','9189460','9189485','9189518','9189538','9189540','9189545','9189551','9189578','9195616','9195627','9219394','9219395','9227370','9227991','9240771','9241490','9249938','9259058','9259782','9324392','9326860','9326950','9326954','9328292','9328296','9328321','9328490','9328920','9328923','9328926','9328933','9333106','9336855','9352394','9352414','9352793','9353114','9353784','9354031','9354484','9362188','9373247','9401471','9411673','9411674','9412380','9412382','9412774','9458900','9458902','9464969','9468224','9468225','9468243','9468893','9469438','9469444','9470273','9470282','9483094','9515051','9520020','9520290','9521704','9523286','9523287','9523293','9523321','9523324','9545683','9588031','9588043','9588112','9588154','9588172','9591097','9611182','9611183','9612047','9622957','9622972','9622994','9623171','9623282','9623359','9623773','9623801','9623810','9623815','9623840','9624746','9633328','9633349','9653906','9656773','9656777','9656787','9671737','9671816','9671984','9702667','9702686','9724338','9726103','9726278','9726300','9726307','9753407','9753477','9753803','9756072','9786210','9786216','9786426','9786454','9786456','9786473','9786523','9786532','9786537','9786621','9786883','9789507','9789609','9800625','9800626','9800830','9800978','9801100','9801107','9801142','9801679','9801971','9801974','9801980','9809324','9809456','9809466','9813166','9813429','9823831','9825828','9826446','9827199','9827264','9827268','9827279','9842380','9844661','9844684','9850715','9858550','9858561','9860190','9872885','9872985','9873009','9873049','9873080','9873111','9873121','9873123','9873132','9873138','9873149','9879485','9879709','9880720','9885502','9893610','9893626','9893627','9893719','9898174','9905943','9918203','9919191','9919197','9921303','9929104','9933981','9933984','9934256','9934261','9934361','9934365','9934382','9934385','9934424','9936787','9936913','9936958','9937072','9937075','9942194','9943644','9946939','9955423','9955426','9955430','9964626','9964628','9966179','9967214','9971397','9986146','9987984','9988234'
)
			)TBL
	GROUP BY GROUP_NO, LOCATION, LOC_NAME)STK ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO
};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_lc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "instock_us.csv: $!";
 
$dbh->disconnect;

}


sub mail1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

#$to = ' cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, jerson.roma@metrogaisano.com, bermon.alcantara@metrogaisano.com, eljie.laquinon@metrogaisano.com, nilynn.yosores@metrogaisano.com, analiza.dano@metrogaisano.com, anafatima.mancho@metrogaisano.com, cindy.yu@metrogaisano.com, emily.silverio@metrogaisano.com, luz.bitang@metrogaisano.com';

#$bcc = 'kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, rex.cabanilla@metrogaisano.com';
$bcc = 'lea.gonzaga@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'In Stock Report, International - US Items';

$msgbody_file = 'message.txt';

$attachment_file = 'INSTOCK_REPORT_INTERNATIONAL_v2_US.xlsx';

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

#$to = ' manuel.degamo@metrogaisano.com, ace.olalia@metrogaisano.com, alma.espino@metrogaisano.com, angeli_christi.ladot@metrogaisano.com, angelito.dublin@metrogaisano.com, arlene.yanson@metrogaisano.com, augosto.daria@metrogaisano.com, charm.buenaventura@metrogaisano.com, teena.velasco@metrogaisano.com, cristy.sy@metrogaisano.com, diana.almagro@metrogaisano.com, edgardo.lim@metrogaisano.com, edris.tarrobal@metrogaisano.com, fidela.villamor@metrogaisano.com, genaro.felisilda@metrogaisano.com, genevive.quinones@metrogaisano.com, glenda.navares@metrogaisano.com, joefrey.camu@metrogaisano.com, jonalyn.diaz@metrogaisano.com ';

#$bcc = 'lea.gonzaga@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'In Stock Report, International - US Items';

$msgbody_file = 'message.txt';

$attachment_file = 'INSTOCK_REPORT_INTERNATIONAL_v2_US.xlsx';

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

#$to = ' josemarie.graciadas@metrogaisano.com, jovany.polancos@metrogaisano.com, judy.gilo@metrogaisano.com, julie.montano@metrogaisano.com, kathlene.procianos@metrogaisano.com, limuel.ulanday@metrogaisano.com, cristina.de_asis@metrogaisano.com, mariajoana.cruz@metrogaisano.com, may.sasedor@metrogaisano.com, michelle.calsada@metrogaisano.com, policarpo.mission@metrogaisano.com, rex.refuerzo@metrogaisano.com, ricky.tulda@metrogaisano.com, ronald.dizon@metrogaisano.com, roselle.agbayani@metrogaisano.com, rowena.tangoan@metrogaisano.com, roy.igot@metrogaisano.com, tessie.cabanero@metrogaisano.com, victoria.ferolino@metrogaisano.com, wendel.gallo@metrogaisano.com, juanjose.sibal@metrogaisano.com, julie.montano@metrogaisano.com ';

#$bcc = 'lea.gonzaga@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'In Stock Report, International - US Items';

$msgbody_file = 'message.txt';

$attachment_file = 'INSTOCK_REPORT_INTERNATIONAL_v2_US.xlsx';

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





