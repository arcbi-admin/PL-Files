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

printf "IN STOCK REPORT \n";

&generate_csv;
&generate_csv_oos;
&generate_csv_nbb;

&zip_file;

&mail1;	
&mail2;	
&mail3;	
&mail4;	

exit;
 
#================================= FUNCTIONS ==================================#

sub generate_csv {

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
# my $uid = 'kent';
# my $pw = 'amer1c8';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $csv2 = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh2, ">", "instock_raw_file_v12.csv" or die "instock_raw_file_v12.csv: $!";

$test2 = qq{
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
SGD.DEPT, SGD.DEPT_NAME,
SUM(STK.TOT_REPL_ITEMS) TOT_REPL_ITEMS, 
SUM(STK.REPL_WITH_SOH) REPL_WITH_SOH
FROM
	(SELECT S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME, DEPT, DEPT_NAME
		FROM
			(SELECT DISTINCT STORE, STORE_NAME FROM STORE WHERE STORE NOT IN (3010,6008) UNION ALL SELECT DISTINCT WH AS STORE, WH_NAME STORE_NAME FROM WH)S,   
			(SELECT B.MERCH_GROUP_CODE, B.MERCH_GROUP_DESC, B.ATTRIB1 ATTRIB1, B.ATTRIB2, G.DIVISION, I.DIV_NAME, D.GROUP_NO, G.GROUP_NAME, D.PURCHASE_TYPE, D.DEPT, D.DEPT_NAME 
				FROM DEPS D 
				  JOIN GROUPS G ON D.GROUP_NO = G.GROUP_NO
				  JOIN DIVISION I ON G.DIVISION = I.DIVISION 
				  JOIN BI_MERCH_GROUP B ON I.DIVISION = B.DIVISION)M
		GROUP BY S.STORE, S.STORE_NAME, M.MERCH_GROUP_CODE, M.MERCH_GROUP_DESC, M.ATTRIB1, M.ATTRIB2, M.DIVISION, M.DIV_NAME, M.GROUP_NO, M.GROUP_NAME, DEPT, DEPT_NAME
	)SGD
	LEFT JOIN
	(SELECT GROUP_NO, DEPT, LOCATION, LOC_NAME, COUNT(REPL_TAG) TOT_REPL_ITEMS, COUNT(STOCK_TAG) REPL_WITH_SOH
	FROM(
		SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, REPL.ITEM, REPL.LOCATION, LOCATIONS.LOC_NAME, 'Y' REPL_TAG, 
		  STOCK_ON_HAND, CASE WHEN STOCK_ON_HAND <= 0 THEN NULL ELSE 'Y' END AS STOCK_TAG --DECODE(STOCK_ON_HAND,0,'','Y') STOCK_TAG --, REPL_B.W_ITEM AUTO_REP
		FROM REPL_ITEM_LOC REPL
			--LEFT JOIN (SELECT DISTINCT ITEM, 'Y' AS W_ITEM FROM REPL_ITEM_LOC WHERE (DEACTIVATE_DATE IS NULL OR DEACTIVATE_DATE > SYSDATE) AND LOC_TYPE = 'W')REPL_B ON REPL.ITEM=REPL_B.ITEM
			LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
			LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
			LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
			LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
			LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
			LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
			LEFT JOIN BI_MERCH_GROUP BI ON D.DIVISION = BI.DIVISION
			LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON REPL.LOCATION=LOCATIONS.LOC
		WHERE LOC.STATUS IN ('A') AND DEPS.PURCHASE_TYPE = 0 AND REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE) 
      AND MST.ORDERABLE_IND = 'Y' AND SOH.FIRST_RECEIVED IS NOT NULL  
      AND (
          (( BI.MERCH_GROUP_CODE='SU' OR ((BI.MERCH_GROUP_CODE='OT' AND D.DIVISION=8500) OR (BI.MERCH_GROUP_CODE='OT' AND D.DIVISION=8000 AND GROUPS.GROUP_NO = 8040))) 
              AND (SYSDATE-SOH.LAST_RECEIVED) <= 91 ) 
            OR
          (( BI.MERCH_GROUP_CODE='DS' OR ((BI.MERCH_GROUP_CODE='OT' AND D.DIVISION=9000) OR (BI.MERCH_GROUP_CODE='OT' AND D.DIVISION=8000 AND GROUPS.GROUP_NO != 8040))) 
              AND (SYSDATE-SOH.LAST_RECEIVED) <= 182 )
          )
	)TBL
	--WHERE AUTO_REP = 'Y'
	GROUP BY GROUP_NO, DEPT, LOCATION, LOC_NAME)STK
	ON SGD.STORE = STK.LOCATION AND SGD.GROUP_NO = STK.GROUP_NO AND SGD.DEPT = STK.DEPT
GROUP BY 
CASE WHEN TO_CHAR(SGD.STORE) = '4002' THEN '2001W' ELSE TO_CHAR(SGD.STORE) END,
CASE WHEN SGD.STORE IN ('2012', '2013', '3009', '4004', '3010', '3011', '3012') THEN 'SU' || SGD.STORE 	 WHEN SGD.STORE IN ('4002') THEN 'SU2001W'     WHEN SGD.STORE = '2223' THEN 'DS' || SGD.STORE	 ELSE SGD.MERCH_GROUP_CODE || SGD.STORE END,	 SGD.STORE_NAME, 
SGD.MERCH_GROUP_CODE, 
CASE WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 9000) THEN 'DS'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8500) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO = 8040) THEN 'SU'     WHEN (SGD.MERCH_GROUP_CODE = 'OT' AND SGD.DIVISION = 8000 AND SGD.GROUP_NO != 8040) THEN 'DS'ELSE SGD.MERCH_GROUP_CODE END,
SGD.MERCH_GROUP_DESC, 
SGD.ATTRIB1, SGD.ATTRIB2, 
SGD.DIVISION, SGD.DIV_NAME, 
SGD.GROUP_NO, SGD.GROUP_NAME,
SGD.DEPT, SGD.DEPT_NAME
ORDER BY 1, 3, 5, 7, 9
}; 
 
my $sth2 = $dbh->prepare ($test2);
 $sth2->execute;
 $csv2->print ($fh2, $sth2->{NAME_lc});
 while (my $row2 = $sth2->fetch) {
     $csv2->print ($fh2, $row2) or $csv2->error_diag;
     }
 close $fh2 or die "instock_raw_file_v12.csv: $!";

 
$dbh->disconnect;

}

sub generate_csv_oos {

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
# my $uid = 'kent';
# my $pw = 'amer1c8';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $csv2 = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh2, ">", "oos_items_raw_file_v12.csv" or die "oos_items_raw_file_v12.csv: $!";

$test2 = qq{
SELECT LOCATION, LOC_NAME, GROUP_NO, GROUP_NAME, DEPT, DEPT_NAME, SHORT_DESC AS ITEM_DESC, ITEM, STATUS, SUP_NAME AS VENDOR_NAME, PRIMARY_SUPP AS VENDOR, 
  STOCK_ON_HAND, RTV_QTY, CUSTOMER_RESV AS CUSTOMER_RESERVE, IN_TRANSIT_QTY AS IN_TRANSIT, (TSF_EXPECTED_QTY + QTY_ALLOCATED) AS INBOUND_ALLOC, 
  ON_ORDER, TSF_RESERVED_QTY, LAST_RECEIVED, LAST_SOLD,  
  NVL( (CASE WHEN LOCATION BETWEEN 7000 AND 7999 OR LOC_TYPE = 'W' THEN OUTOFSTOCKDATEWHSE ELSE OUTOFSTOCKDATESTORE END),CREATE_DATETIME) AS  OUT_OF_STOCK, 
  SALES_ISSUES AS ISSUANCES_AVG_12_WKS
FROM (
SELECT DISTINCT REPL.LOCATION, LO.LOC_NAME, SOH.LOC_TYPE, GROUPS.GROUP_NO, GROUPS.GROUP_NAME, DEPS.DEPT, DEPS.DEPT_NAME, MST.SHORT_DESC, REPL.ITEM, LOC.STATUS, SUPS.SUP_NAME, 
	SOH.PRIMARY_SUPP, SOH.STOCK_ON_HAND, SOH.RTV_QTY, SOH.CUSTOMER_RESV, SOH.IN_TRANSIT_QTY, SOH.TSF_EXPECTED_QTY, APPRUNSHIP.QTY_ALLOCATED, ONORD.ON_ORDER, 
	SOH.TSF_RESERVED_QTY, SOH.LAST_RECEIVED, SOH.LAST_SOLD AS LAST_SOLD, SOH.CREATE_DATETIME, SOH.LAST_SOLD AS OUTOFSTOCKDATESTORE, 
	SHP.OUTOFSTOCKDATEWHSE, SLS.SALES_ISSUES 
FROM REPL_ITEM_LOC REPL
	LEFT JOIN ITEM_LOC LOC ON REPL.ITEM=LOC.ITEM AND REPL.LOCATION=LOC.LOC
	LEFT JOIN ITEM_LOC_SOH SOH ON REPL.ITEM=SOH.ITEM AND REPL.LOCATION=SOH.LOC
	LEFT JOIN ITEM_MASTER MST ON REPL.ITEM=MST.ITEM 
	LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
	LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
	LEFT JOIN DIVISION D ON GROUPS.DIVISION = D.DIVISION
	LEFT JOIN BI_MERCH_GROUP BI ON D.DIVISION = BI.DIVISION
	LEFT JOIN SUPS ON SUPS.SUPPLIER = SOH.PRIMARY_SUPP
	LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LO ON REPL.LOCATION=LO.LOC
	LEFT JOIN (SELECT TO_LOC, ITEM, SUM(NVL(QTY_ALLOCATED,0))QTY_ALLOCATED 
				 FROM ALLOC_HEADER HD 
					INNER JOIN ALLOC_DETAIL DTL ON HD.ALLOC_NO=DTL.ALLOC_NO 
				 WHERE HD.STATUS='A' AND (QTY_CANCELLED IS NULL OR QTY_CANCELLED=0) AND QTY_TRANSFERRED=0 
				 GROUP BY TO_LOC, ITEM)APPRUNSHIP ON REPL.ITEM = APPRUNSHIP.ITEM AND REPL.LOCATION = APPRUNSHIP.TO_LOC
	LEFT JOIN (SELECT HD.LOCATION, LOC.ITEM, SUM(NVL(QTY_ORDERED,0)-NVL(QTY_RECEIVED,0)) ON_ORDER 
				 FROM ORDHEAD HD 
					INNER JOIN ORDLOC LOC ON HD.ORDER_NO=LOC.ORDER_NO 
				 WHERE HD.STATUS='A' AND QTY_ORDERED<>0 GROUP BY HD.LOCATION, LOC.ITEM)ONORD ON REPL.ITEM = ONORD.ITEM AND REPL.LOCATION = ONORD.LOCATION
	LEFT JOIN (SELECT LOC, ITEM, AVG(SALES_ISSUES) SALES_ISSUES
				FROM ITEM_LOC_HIST WHERE EOW_DATE > SYSDATE-84
				GROUP BY LOC, ITEM)SLS ON SLS.LOC = REPL.LOCATION AND SLS.ITEM = REPL.ITEM
	LEFT JOIN (SELECT S.FROM_LOC, SK.ITEM, MAX(SHIP_DATE) AS OUTOFSTOCKDATEWHSE
				FROM SHIPMENT S, SHIPSKU SK WHERE S.SHIPMENT = SK.SHIPMENT
				GROUP BY S.FROM_LOC, ITEM)SHP ON SHP.FROM_LOC = REPL.LOCATION AND SHP.ITEM = REPL.ITEM
WHERE LOC.STATUS IN ('A') AND DEPS.PURCHASE_TYPE = 0 AND SOH.STOCK_ON_HAND <= 0 AND REPL.ACTIVATE_DATE <= SYSDATE AND (REPL.DEACTIVATE_DATE IS NULL OR REPL.DEACTIVATE_DATE > SYSDATE) 
	AND MST.ORDERABLE_IND = 'Y' AND SOH.FIRST_RECEIVED IS NOT NULL
	AND (
          (( BI.MERCH_GROUP_CODE='SU' OR ((BI.MERCH_GROUP_CODE='OT' AND D.DIVISION=8500) OR (BI.MERCH_GROUP_CODE='OT' AND D.DIVISION=8000 AND GROUPS.GROUP_NO = 8040))) 
              AND (SYSDATE-SOH.LAST_RECEIVED) <= 91 ) 
            OR
          (( BI.MERCH_GROUP_CODE='DS' OR ((BI.MERCH_GROUP_CODE='OT' AND D.DIVISION=9000) OR (BI.MERCH_GROUP_CODE='OT' AND D.DIVISION=8000 AND GROUPS.GROUP_NO != 8040))) 
              AND (SYSDATE-SOH.LAST_RECEIVED) <= 182 )
          )
)
}; 
 
my $sth2 = $dbh->prepare ($test2);
 $sth2->execute;
 $csv2->print ($fh2, $sth2->{NAME_lc});
 while (my $row2 = $sth2->fetch) {
     $csv2->print ($fh2, $row2) or $csv2->error_diag;
     }
 close $fh2 or die "oos_items_raw_file_v12.csv: $!";

 
$dbh->disconnect;

}

sub generate_csv_nbb {

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
# my $uid = 'kent';
# my $pw = 'amer1c8';
my $uid = 'rmsprd';
my $pw = 'noida123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw) or die "Unable to connect: $DBI::errstr";

my $csv2 = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh2, ">", "nbb_raw_file.csv" or die "nbb_raw_file.csv: $!";

$test2 = qq{
SELECT LOCATION, LOC_NAME, GROUP_NO, DEPT, ITEM, SHORT_DESC, STATUS, STOCK_ON_HAND, RTV_QTY, CUSTOMER_RESV, IN_TRANSIT_QTY, TSF_EXPECTED_QTY, QTY_ALLOCATED, ON_ORDER, TSF_RESERVED_QTY, LAST_RECEIVED, LAST_SOLD, CREATE_DATETIME
FROM(
	SELECT DISTINCT DEPS.GROUP_NO, MST.DEPT, DEPT_NAME, SOH.ITEM, MST.SHORT_DESC, SOH.LOC LOCATION, LOCATIONS.LOC_NAME, LOC.STATUS, 
	  SOH.STOCK_ON_HAND, SOH.RTV_QTY, SOH.CUSTOMER_RESV, SOH.IN_TRANSIT_QTY, SOH.TSF_EXPECTED_QTY, APPRUNSHIP.QTY_ALLOCATED, ONORD.ON_ORDER, SOH.TSF_RESERVED_QTY, SOH.LAST_RECEIVED, SOH.LAST_SOLD AS LAST_SOLD, SOH.CREATE_DATETIME
	FROM ITEM_LOC_SOH SOH
		LEFT JOIN ITEM_LOC LOC ON SOH.ITEM=LOC.ITEM AND SOH.LOC=LOC.LOC
		LEFT JOIN ITEM_MASTER MST ON SOH.ITEM=MST.ITEM 
		LEFT JOIN DEPS ON MST.DEPT=DEPS.DEPT
		LEFT JOIN GROUPS ON DEPS.GROUP_NO=GROUPS.GROUP_NO
		LEFT JOIN (SELECT DISTINCT STORE LOC, STORE_NAME LOC_NAME FROM STORE UNION ALL SELECT DISTINCT WH AS LOC, WH_NAME LOC_NAME FROM WH)LOCATIONS ON SOH.LOC=LOCATIONS.LOC
    LEFT JOIN (SELECT TO_LOC, ITEM, SUM(NVL(QTY_ALLOCATED,0))QTY_ALLOCATED 
				 FROM ALLOC_HEADER HD 
					INNER JOIN ALLOC_DETAIL DTL ON HD.ALLOC_NO=DTL.ALLOC_NO 
				 WHERE HD.STATUS='A' AND (QTY_CANCELLED IS NULL OR QTY_CANCELLED=0) AND QTY_TRANSFERRED=0 
				 GROUP BY TO_LOC, ITEM)APPRUNSHIP ON SOH.ITEM = APPRUNSHIP.ITEM AND SOH.LOC = APPRUNSHIP.TO_LOC
    LEFT JOIN (SELECT HD.LOCATION, LOC.ITEM, SUM(NVL(QTY_ORDERED,0)-NVL(QTY_RECEIVED,0)) ON_ORDER 
				 FROM ORDHEAD HD 
					INNER JOIN ORDLOC LOC ON HD.ORDER_NO=LOC.ORDER_NO 
				 WHERE HD.STATUS='A' AND QTY_ORDERED<>0 GROUP BY HD.LOCATION, LOC.ITEM)ONORD ON SOH.ITEM = ONORD.ITEM AND SOH.LOC = ONORD.LOCATION      
	WHERE LOC.STATUS IN ('A')  AND 
	(
	(SOH.LOC IN ( '2004','2005','2006','2007','2009','2012','2013','2223','3012','4004','6002','6012' ) AND SOH.ITEM IN ( '2070582','2006802','10483236','8458735','8201966','2098999','7863875','3758472','2099064','5239818','2099125','2035000','8956972','8956989','2003603','2035314','1987089','1994162','2003429','2003528','2003436','2003450','10095188','10062286','9083975','2071664','9083978','9834562','2071626','9617526','2071671','2071619','9083977','2025223','9230619','2070209','9516657','3118603','2063737','2094106','5544264','3118610','4082392','2104485','5244904','8404954','3118634','2046563','6256548','9192574','2112411','2077017','9245716','2112459','10516960','10343097','10573829','2014173','2020839','1992885','10302110','10071291','10627605','101590325','2089607','9604234','10627604','1999785','9018150','9612599','9985068','9287582','9985071','2088853','15003212','9058949','9420841','10666074','9790857','9945757','9628055','3794753','2069180','8779199','9945758','9612426','15003003','6948405','3794852','2101132','9997581','2065519','9506663','7173233','3762387','8264954','10153202','7791413','9413022','10075820','10154458','10154459','9798593','9413018','9128159','9798596','9413021','9304829','9531038','2002446','10154461','9531053','2002941','5788293','9781681','9577940','10506394','10005413','15003006','9756146','9756145','2101699','9732968','9732967','2099392','9229170','10225270','9495258','10036303','8417480','3313855','8846051','8846105','3313886','9097724','3322376','9352311','9198427','10220009','9802970','9414472','9414467','10044690','2059945','8403988','2047683','9166472','9111641','9111655','2106441','10163810','10163811','10730252','9354893','10331962','9378636','10517402','9378666','10331991','10323783','9622185','9378645','9409678','10573834','10144094','8670106','9059889','8669797','10217611','9099661','9099310','2094670','10249301','10144084','2077239','10181461','6447779','8404794','2088860','2088792','15003036','9616898','6393335','2068268','6393366','6171261','3758632','2083223','2068336','10739487','2068305','2083278','9288596','101590122','10677353','6768300','2083247','9303560','10091953','2068275','6171278','2068299','2061559','2061542','6393373','10102570','9600964','9768446','101563256','9351396','2068282','9053954','10571068','101580575','9768448','10405061','9229174','9879223','2097497','9628042','9592783','2099224','9530923','2108650','9592734','2042831','2097473','9337737','9592793','10151283','2089409','2099217','2089430','10156677','101535463','10768217','2089386','9285894','3203873','10082124','10082128','2091693','10228601','10409555','2060347','1991680','8369840','2035604','2119489','9257330','1991758','2059099','2036311','9023384','9397210','9315386','5402625','9128577','3772331','2060330','3772997','2978185','8432315','10268141','9057888','9612608','9304841','9482243','9912519','10005140','9719678','15001915','15001916','2043739','2053592','2002064','5646265','9229318','8807687','9938810','2040219','10752794','10640020','10121636','10752795','10264849','7225017','9609446','10205309','9609452','10205298','4.80889E+12','2102139','2101996','9653550','3786420','3356647','10779744','9354924','3356630','3356654','3311325','3311332','9741484','9035311','9741446','9741440','10673057','6828127','3972','8055','3971','8164','2101378','3792841','3792414','9129079','9403659','2101385','9111102','9192668','10151714','10151719','2070230','10224003','10010001','2070247','2070261','10224013','9495208','9270797','10168607','9432399','9215252','9495264','9495230','9375116','9375069','9291066','9375127','15002201','15002211','1992977','3132173','2052052','3132166','2051963','2088914','8202529','8713155','15002228','9832635','9523786','9261923','9847335','15000321','2095929','9597891','10212653','1990928','9592985','9788570','9788582','9985192','9116097','10066546','9716380','1840223','15003011','9462912','9981874','9897461','2004310','2004327','9034647','9034649','2004334','9780086','101422941','10082044','10174443','10174453','10196649','9765501','10022702','9292448','15002939','9850872','9893594','9893598','9507012','2096667','9507016','2096407','10430482','10430481','9837032','10195624','10430333','10179705','9894887','10430315','2033839','9033679','2116273','2115597','2450070','9552315','8947215','10200258','10774383','10285304','10285308','10285307','9853613','9853597','9853623','9358632','3365281','9787751','9787778','9787764','9788271','10252423','9787779','10242216','10017759','10017901','10020873','101574326','9963896','10701126','101666075','101574401','9917884','10734466','10122103','10075303','10085788','10456639','10507565','9308960','10612338','10629076','9887384','10178007','9377775','9377770','10508099','10752128','10178012','2086453','8418579','7592775','10002248','9497411','7592751','9948856','10215077','1999778','9890643','9380279','9341571','9351111','9920013','101477584','9358673','9726607','10746426','10567209' )) OR 
	(SOH.LOC IN ( '2001','2002','2003','2008','2010','2011','3001','3002','3003','3004','3005','3006','3007','3009','4002','4003','6001','6003','6004','6005','6009','6010' ) AND SOH.ITEM IN ( '2070582','10112680','2006802','8458735','8201966','2098999','3344729','3316016','2099064','2099125','5239818','2096742','2023397','2080079','5263073','2978154','8807441','2003320','2064932','6545802','2003337','2067575','2067582','8956972','2035000','2116518','2116525','8956989','2035314','2013763','2003603','10062276','9083975','2071664','9083978','2071626','2071671','9083977','9834562','2070209','9516657','9457103','3118603','8404954','5544264','2094106','9192574','2112411','6256548','2063737','9245716','9245724','2104485','2112459','3118634','9026696','10203379','2112428','3118610','2019116','2046563','2077017','5244904','9967141','9543064','9926921','9285449','2016269','9967126','2112442','10605583','10083470','10083528','10083485','10516960','10343097','10573829','9612895','2020839','2014173','9118033','101657019','101433076','9544940','10730877','10248486','101657127','2085241','9608501','2085104','2085289','10302110','10071291','9604234','10627605','10627604','101590325','2089607','10750783','1992236','9018150','2058894','2100760','15003212','9287582','3781357','9166462','2110912','2088853','9985068','9058949','9279558','9985071','10666074','7530982','9420841','9043768','9341489','7385018','9025360','2025575','9509344','9619829','2061702','9893963','9118909','9340661','10031768','3794753','8779199','9628055','9128164','9945757','9612426','2065502','9172002','9268195','2069180','9169114','9790873','2101262','2065519','2001494','7173233','5272112','2089126','2088785','9104215','8402103','2065205','3762387','8840080','7791413','9413022','10506394','9531038','10154458','10075820','9413018','10154459','9413021','3322666','3320747','9101411','3316092','7114359','10077213','2101699','10077203','2099392','9229170','2089867','8417480','2030722','10036303','9070021','9070024','3313886','3334850','3334775','3313855','3322376','7081873','7081897','3313831','7057779','9905812','9988438','9802970','9580774','9414468','10044690','8403988','6023751','2059976','9257768','9111641','9111655','10163810','10163811','10331962','10517402','9378636','9378666','10574018','2094670','8669797','2055701','10520342','8670106','9117268','2055718','2077239','10242404','9057463','10144094','10242448','9034675','9809571','10217611','8765901','9099657','10657317','9000449','10249307','8404794','2088792','2088860','6447779','15003036','9219375','2068510','5672202','9116770','7813399','6768300','6393335','2068268','9013457','2068336','2061542','2083223','9506226','2061559','6393366','9351396','10135507','9879223','9229174','2099224','2089416','2089386','10156677','2099217','2099194','2089409','2112701','10082124','2095424','2112640','2112657','2112718','10113551','10082128','2091693','10228601','10409555','2089881','2092904','2092928','9142746','2099897','9142740','2092935','9236135','10124877','2092911','2978185','9257330','2035604','1991680','3772997','5402625','1991758','2060347','9038253','2061597','2065410','9255705','2065403','5226344','6240936','9438033','5845361','9719678','15001916','9799862','2043739','2002040','2002064','101659352','101569507','9938810','2040219','10752794','10264849','10121636','2102139','3786420','3355893','9354924','8429001','10779744','3321508','9354914','3311189','3311639','3356654','3356647','9741440','6828127','3405','8003','3610','8055','3959','101651594','3792841','15003437','8938060','9495264','9432394','2070230','8417473','9354884','9066871','9354578','9066867','9128167','2061900','9777187','15002211','2099767','2060958','2008837','2051963','2088914','8202529','8202512','9796444','2116273','9663718','9552315','9083071','101703219','101703159','9451161','10137016','15002230','15002229','9261923','9523786','10641595','2095929','10031647','10031646','10212653','9592985','9499921','2091846','9788582','9788570','9852058','9716380','9678254','9243002','10001076','9462912','9889533','9897461','9981874','9897463','9897462','2004310','2004327','9812892','9629308','9994625','10763769','9871032','10003967','10082044','9780086','10003963','10135767','9457562','9292448','9363194','9292447','10022702','9617966','10690113','9850872','9893594','9507012','2096667','2096407','8202055','10216301','10430482','10737453','10195624','9837032','10179705','10334147','10385248','10385290','10430333','9894887','3323786','3323755','10285307','10285308','10285302','10122798','10122796','2058399','10285266','2006628','9118164','9118198','9571818','3365281','9571821','2072531','9163010','2120003','2071909','2071916','9425602','3359662','2091464','3360224','9787751','2067544','9788271','2085845','9787764','2085838','10242216','9787778','9788272','10774384','10017759','10017901','10020873','10456639','101574326','9215824','101574401','10629075','9660462','10629076','9904117','101666075','10701126','10690146','10178007','9887384','9377775','10178012','2086453','2086460','2086446','9777626','8468666','8418579','7592751','7592775','7592829','7592805','10165789','10165797','10165779','10165782','9183620','9183622','8904829','8904843','9530162','15000381','10215077','1999778','9380279','2025339','9287404','2031590','9755384','9920013','9258104' )) 
	)
) ORDER BY 1, 3, 4, 5
}; 
 
my $sth2 = $dbh->prepare ($test2);
 $sth2->execute;
 $csv2->print ($fh2, $sth2->{NAME_lc});
 while (my $row2 = $sth2->fetch) {
     $csv2->print ($fh2, $row2) or $csv2->error_diag;
     }
 close $fh2 or die "nbb_raw_file.csv: $!";

 
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
my $file_member2 = $zip->addFile( 'oos_items_raw_file_v12.csv' );
my $file_member3 = $zip->addFile( 'nbb_raw_file.csv' );

# save the zip file
unless ( $zip->writeToFileNamed('Replenishment In-stock – SKU Details.rar') == AZ_OK ) {
	die 'write error';
}
   
}


sub mail1 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file, $attachment_file_2 ) = @ARGV;

$to = ' fili.mercado@metrogaisano.com, emily.silverio@metrogaisano.com, ronald.dizon@metrogaisano.com, chit.lazaro@metrogaisano.com, jocelyn.sarmiento@metrogaisano.com, charisse.mancao@metrogaisano.com, cindy.yu@metrogaisano.com, cresilda.dehayco@metrogaisano.com, evan.inocencio@metrogaisano.com, fe.botero@metrogaisano.com, jonrel.nacor@metrogaisano.com, junah.oliveron@metrogaisano.com, lyn.cabatuan@metrogaisano.com, zenda.mangabon@metrogaisano.com, joyce.mirabueno@metrogaisano.com, mariegrace.ong@metrogaisano.com, cherry.gulloy@metrogaisano.com, janice.bedrijo@metrogaisano.com, jerson.roma@metrogaisano.com, bermon.alcantara@metrogaisano.com, nilynn.yosores@metrogaisano.com, anafatima.mancho@metrogaisano.com, leslie.chipeco@metrogaisano.com, karan.malani@metrogaisano.com, may.sasedor@metrogaisano.com, donna.fernando@metrogaisano.com, victoria.ferolino@metrogaisano.com, wendel.gallo@metrogaisano.com, alexander.tejedor@metrogaisano.com, judith.tud@metrogaisano.com, marygrace.ong@metrogaisano.com, rowena.tangoan@metrogaisano.com, genevive.quinones@metrogaisano.com, angeli_cristi.ladot@metrogaisano.com, edris.tarrobal@metrogaisano.com, charm.buenaventura@metrogaisano.com, augosto.daria@metrogaisano.com, cj.jesena@metrogaisano.com, gerry.guanlao@metrogaisano.com, opcplanning@metrogaisano.com ';

$cc = 'annalyn.conde@metrogaisano.com, lea.gonzaga@metrogaisano.com, fnaquines@metro.com.ph, cham.burgos@metrogaisano.com, kent.mamalias@metrogaisano.com, rlegaspi@metro.com.ph';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';	

$subject = 'Replenishment In-stock – SKU Details - [In Stock, NBB, OOS]';

$msgbody_file = 'message.txt';

#$attachment_file = "Replenishment In-stock v1.21.xlsx";
$attachment_file_2 = "Replenishment In-stock – SKU Details.rar";

my $msgbody = read_file( $msgbody_file );

#my $attachment_data = encode_base64( read_file( $attachment_file, 1 ));
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
Content-Type: application/octet-stream; name="$attachment_file_2"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$attachment_file_2"
$attachment_data_2
$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

}

sub mail2 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file, $attachment_file_2 ) = @ARGV;

$to = ' noli.lee@metrogaisano.com, jovany.polancos@metrogaisano.com, alma.espino@metrogaisano.com, arlene.yanson@metrogaisano.com, josemarie.graciadas@metrogaisano.com, ronald.parragatos@metrogaisano.com, vivian.ablang@metrogaisano.com, emma.villoson@metrogaisano.com, rachel.riva@metrogaisano.com, roselle.agbayani@metrogaisano.com, michelle.calsada@metrogaisano.com, jonalyn.diaz@metrogaisano.com, joseph.landicho@metrogaisano.com, rex.refuerso@metrogaisano.com, al_rey.candia@metrogaisano.com, cheruvim.villaceran@metrogaisano.com, mae_flor.lauronal@metrogaisano.com, egan.pacquiao@metro.com.ph, evanguardia@metro.com.ph, rysalas@metro.com.ph, jdorimon@metro.com.ph, eestrera@metro.com.ph, efermo@metro.com.ph, irene.montemayor@metrogaisano.com, limuel.ulanday@metrogaisano.com, joefrey.camu@metrogaisano.com, karan.malani@metrogaisano.com, lia.chipeco@metrogaisano.com, danice.tanael@metrogaisano.com ';

$cc = 'kent.mamalias@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';	

$subject = 'Replenishment In-stock – SKU Details - [In Stock, NBB, OOS]';

$msgbody_file = 'message.txt';

#$attachment_file = "Replenishment In-stock v1.21.xlsx";
$attachment_file_2 = "Replenishment In-stock – SKU Details.rar";

my $msgbody = read_file( $msgbody_file );

#my $attachment_data = encode_base64( read_file( $attachment_file, 1 ));
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
Content-Type: application/octet-stream; name="$attachment_file_2"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$attachment_file_2"
$attachment_data_2
$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

}

sub mail3 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file, $attachment_file_2 ) = @ARGV;

$to = ' carmelita.intia@metrogaisano.com, consorcia.mullon@metrogaisano.com, contessa.fernandez@metrogaisano.com, delia.jakosalem@metrogaisano.com, editha.cabriles@metrogaisano.com, edna.prieto@metrogaisano.com, jecil.cumayas@metrogaisano.com, jennifer.yu@metrogaisano.com, lorena.madraga@metrogaisano.com, maryann.delarama@metrogaisano.com, maryjoy.montes@metrogaisano.com, mecelle.quimbo@metrogaisano.com, mirasol.barcoma@metrogaisano.com, nenita.cabigon@metrogaisano.com, teresita.manatad@metrogaisano.com, vilma.paner@metrogaisano.com, melinda.uy@metrogaisano.com, jgeniston@metro.com.ph, cartes@metro.com.ph, hcaberte@metro.com.ph, mvilla@metro.com.ph, mcabungcal@metro.com.ph, mcombinido@metro.com.ph, april.agapito@metrogaisano.com, jordan.mok@metrogaisano.com, marlita.portes@metrogaisano.com, advento.resma@metrogaisano.com, arlene.te@metrogaisano.com, armando.pitogo@metrogaisano.com, christine.lanohan@metrogaisano.com, delvie.pitogo@metrogaisano.com, emillie.ponsica@metrogaisano.com ';

$cc = 'kent.mamalias@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';	

$subject = 'Replenishment In-stock – SKU Details - [In Stock, NBB, OOS]';

$msgbody_file = 'message.txt';

#$attachment_file = "Replenishment In-stock v1.21.xlsx";
$attachment_file_2 = "Replenishment In-stock – SKU Details.rar";

my $msgbody = read_file( $msgbody_file );

#my $attachment_data = encode_base64( read_file( $attachment_file, 1 ));
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
Content-Type: application/octet-stream; name="$attachment_file_2"
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="$attachment_file_2"
$attachment_data_2
$boundary--
END_OF_BODY

sendmail(%mail) or die $Mail::Sendmail::error;

print "Sendmail Log says:\n$Mail::Sendmail::log\n";

}

sub mail4 {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file, $attachment_file_2 ) = @ARGV;

$to = ' manuel.degamo@metrogaisano.com, jennifer.moreno@metrogaisano.com, jessica.gaisano@metrogaisano.com, keith.poblete@metrogaisano.com, micah.alvarado@metrogaisano.com, michelle.someros@metrogaisano.com, peachy.aquino@metrogaisano.com, ryan.uson@metrogaisano.com, sheen.ducay@metrogaisano.com, annie.desuyo@metrogaisano.com, marlit.ignacio@metrogaisano.com, mildred.quinones@metrogaisano.com, rosemarie.saravia@metrogaisano.com, rowena.conde@metrogaisano.com, tessie.baldezamo@metrogaisano.com, diana.almagro@metrogaisano.com, angelito.dublin@metrogaisano.com, arlene.yanson@metrogaisano.com, charm.buenaventura@metrogaisano.com, teena.velasco@metrogaisano.com, cristy.sy@metrogaisano.com, fidela.villamor@metrogaisano.com, glenda.navares@metrogaisano.com, jonalyn.diaz@metrogaisano.com, josemarie.graciadas@metrogaisano.com, judy.gilo@metrogaisano.com, cristina.de_asis@metrogaisano.com, mariajoana.cruz@metrogaisano.com, ricky.tulda@metrogaisano.com, roy.igot@metrogaisano.com, analiza.dano@metrogaisano.com, luz.bitang@metrogaisano.com, ricky.aguas@metrogaisano.com, maricel.mayorga@metrogaisano.com, mirasol.barcoma@metrogaisano.com, chedie.lim@metrogaisano.com ';

$cc = 'kent.mamalias@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';	

$subject = 'Replenishment In-stock – SKU Details - [In Stock, NBB, OOS]';

$msgbody_file = 'message.txt';

#$attachment_file = "Replenishment In-stock v1.21.xlsx";
$attachment_file_2 = "Replenishment In-stock – SKU Details.rar";

my $msgbody = read_file( $msgbody_file );

#my $attachment_data = encode_base64( read_file( $attachment_file, 1 ));
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








