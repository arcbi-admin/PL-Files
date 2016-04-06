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


$test_query = qq{ SELECT CASE WHEN EXISTS (SELECT *
					FROM ADMIN_ETL_LOG 
					WHERE TO_DATE(LOG_DATE, 'DD-MON-YY') = TO_DATE(SYSDATE,'DD-MON-YY') AND TASK_ID = 'IntSalesDspTy' AND ERR_CODE = 0) THEN 1 ELSE 0 END STATUS 
					FROM DUAL };

$tst_query = $dbh->prepare($test_query);
$tst_query->execute();

while ( my $x =  $tst_query->fetchrow_hashref()){
	$test = $x->{STATUS};
} 

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
		$as_of = $x->{DATE_FLD};
	}

	printf "Test ETL Status = ". $test ." \nArc BI Sales Performance Part 1\nGenerating Data from Source \n";
 
	$workbook = Excel::Writer::XLSX->new("ARC BI VALIDATION RMS TO ARC - Summary (as of $as_of).xlsx");
	$border1 = $workbook->add_format( border => 1, align => 'center', valign => 'right', size => 10, num_format => 4 );
	$desc = $workbook->add_format( border => 1, align => 'left', valign => 'vcenter', size => 10 );
	$desc2 = $workbook->add_format( border => 1, align => 'center', valign => 'vcenter', text_wrap =>1, size => 10, shrink => 1 );
	
	&sale_header;
	&sale_line;
	&sale_invc_tender;
	&sale_invc_discount;
	&sale_line_discount;
	
	$workbook->close();
	$tst_query->finish();
	$dbh->disconnect; 
	
	&mail;
	
	exit;
	
}

else{
	print "Test Status = ". $test ." \nETL still running\n";
	
	$tst_query->finish();
	
	sleep(600);
	
	goto START;
}
 
sub sale_header{
my $worksheet = $workbook->add_worksheet("sale_header");
	$worksheet->set_column( 0, 0, 5 );
	$worksheet->set_column( 1, 4, 12 );
	$worksheet->set_column( 5, 5, 1 );
	$worksheet->set_column( 6, 9, 12 );
	$worksheet->set_column( 10, 10, 1 );
	$worksheet->set_column( 11, 14, 10 );
	my $a = 2;
	my $col = 0;
	
	foreach my $i ( "STORE", "RMS_SALE_TOT_QTY", "RMS_SALE_NET_VAL", "RMS_SALE_TOT_TAX_VAL", "RMS_SALE_TOT_DISC_VAL", "", "ARC_SALE_TOT_QTY", "ARC_SALE_NET_VAL", "ARC_SALE_TOT_TAX_VAL", "ARC_SALE_TOT_DISC_VAL", "", "DIFF_SALE_TOT_QTY", "DIFF_SALE_NET_VAL", "DIFF_SALE_TOT_TAX_VAL", "DIFF_SALE_TOT_DISC_VAL" ) {
		$worksheet->write( 1, $col++, $i, $desc2 );
	}
	
	my $query = qq {
		SELECT RMS.STORE_CODE, RMS_SALE_TOT_QTY, RMS_SALE_NET_VAL, RMS_SALE_TOT_TAX_VAL, RMS_SALE_TOT_DISC_VAL, ARC.ARC_SALE_TOT_QTY, ARC_SALE_NET_VAL, ARC_SALE_TOT_TAX_VAL, ARC_SALE_TOT_DISC_VAL, (NVL(RMS_SALE_TOT_QTY,0))-(NVL(ARC_SALE_TOT_QTY,0)) DIFF_SALE_TOT_QTY, (NVL(RMS_SALE_NET_VAL,0))-(NVL(ARC_SALE_NET_VAL,0)) DIFF_SALE_NET_VAL, (NVL(RMS_SALE_TOT_TAX_VAL,0))-(NVL(ARC_SALE_TOT_TAX_VAL,0)) DIFF_SALE_TOT_TAX_VAL, (NVL(RMS_SALE_TOT_DISC_VAL,0))-(NVL(ARC_SALE_TOT_DISC_VAL,0)) DIFF_SALE_TOT_DISC_VAL
FROM
 (SELECT STORE_CODE,SUM(SALE_TOT_QTY) AS RMS_SALE_TOT_QTY ,SUM(SALE_NET_VAL) AS RMS_SALE_NET_VAL, SUM(SALE_TOT_TAX_VAL) AS  RMS_SALE_TOT_TAX_VAL, SUM(SALE_TOT_DISC_VAL) AS RMS_SALE_TOT_DISC_VAL
			FROM(
				SELECT SH.TRAN_DATETIME AS TRANS_DATE,SH.TRAN_SEQ_NO AS INVOICE_NO,SH.TRAN_TYPE AS SALE_INVC_TYPE,SH.SUB_TRAN_TYPE AS SALE_INVC_SUB_TYPE,
					SH.STATUS    AS SALE_INVC_STATUS,SH.STORE AS STORE_CODE,
					CASE WHEN LENGTH(SH.TRAN_NO) >=10 THEN SUBSTR(SH.TRAN_NO,1,3) ELSE SUBSTR(LPAD(SH.TRAN_NO,10,'0'),1,3) END AS TILL_CODE,
					SH.CASHIER AS CASHIER_CODE,SH.UPDATE_DATETIME AS UPDATES,SUM(SL.QTY) AS SALE_TOT_QTY,SUM(SL.QTY * SL.UNIT_RETAIL) AS SALE_NET_VAL,
					SUM(SGT.TOTAL_IGTAX_AMT) AS SALE_TOT_TAX_VAL,SUM(SD.SALE_TOT_DISC_VAL) AS SALE_TOT_DISC_VAL ,COUNT(SL.ITEM_SEQ_NO) AS SALE_TOT_ITEM_COUNT,
					0 AS SALE_TOT_PACK_COUNT
						FROM SA_TRAN_HEAD@RMS SH JOIN SA_TRAN_ITEM@RMS SL ON SH.TRAN_SEQ_NO = SL.TRAN_SEQ_NO AND SH.STORE = SL.STORE
								LEFT JOIN SA_TRAN_IGTAX@RMS SGT ON SH.TRAN_SEQ_NO = SGT.TRAN_SEQ_NO AND SL.ITEM_SEQ_NO=SGT.ITEM_SEQ_NO
									LEFT JOIN (SELECT TRAN_SEQ_NO,ITEM_SEQ_NO, SUM(QTY*UNIT_DISCOUNT_AMT) AS SALE_TOT_DISC_VAL 
											FROM  SA_TRAN_DISC@RMS GROUP BY TRAN_SEQ_NO,ITEM_SEQ_NO )  SD
											ON SH.TRAN_SEQ_NO = SD.TRAN_SEQ_NO AND SL.ITEM_SEQ_NO =SD.ITEM_SEQ_NO
					GROUP BY SH.TRAN_DATETIME,SH.TRAN_SEQ_NO,SH.TRAN_TYPE,SH.SUB_TRAN_TYPE,SH.STATUS, SH.STORE,SH.TRAN_NO,SH.CASHIER,SH.UPDATE_DATETIME)F 
					WHERE TRUNC(UPDATES)=(SELECT TO_CHAR(SYSDATE-1,'DD MON YYYY') FROM DUAL) --'02 OCT 2014'
					GROUP BY STORE_CODE)RMS
LEFT JOIN
(SELECT STORE_CODE,SUM(SALE_TOT_QTY) AS ARC_SALE_TOT_QTY ,SUM(SALE_NET_VAL) AS ARC_SALE_NET_VAL, SUM(SALE_TOT_TAX_VAL) AS  ARC_SALE_TOT_TAX_VAL, SUM(SALE_TOT_DISC_VAL) AS ARC_SALE_TOT_DISC_VAL
		FROM(
			SELECT SH.TRAN_DATETIME AS TRANS_DATE,SH.TRAN_SEQ_NO AS INVOICE_NO,SH.TRAN_TYPE AS SALE_INVC_TYPE,SH.SUB_TRAN_TYPE AS SALE_INVC_SUB_TYPE,
				SH.STATUS    AS SALE_INVC_STATUS,SH.STORE AS STORE_CODE,CASE WHEN LENGTH(SH.TRAN_NO) >=10 THEN SUBSTR(SH.TRAN_NO,1,3) 
				ELSE SUBSTR(LPAD(SH.TRAN_NO,10,'0'),1,3) END AS TILL_CODE,SH.CASHIER AS CASHIER_CODE,SH.UPDATE_DATETIME AS UPDATES,
				SUM(SL.QTY) AS SALE_TOT_QTY,SUM(SL.QTY * SL.UNIT_RETAIL) AS SALE_NET_VAL,SUM(SGT.TOTAL_IGTAX_AMT) AS SALE_TOT_TAX_VAL,
				SUM(SD.SALE_TOT_DISC_VAL) AS SALE_TOT_DISC_VAL ,COUNT(SL.ITEM_SEQ_NO) AS SALE_TOT_ITEM_COUNT,
				0 AS SALE_TOT_PACK_COUNT FROM SA_TRAN_HEAD SH JOIN SA_TRAN_ITEM SL ON SH.TRAN_SEQ_NO = SL.TRAN_SEQ_NO
				AND SH.STORE = SL.STORE LEFT JOIN SA_TRAN_IGTAX SGT
				ON SH.TRAN_SEQ_NO = SGT.TRAN_SEQ_NO AND SL.ITEM_SEQ_NO=SGT.ITEM_SEQ_NO
				LEFT JOIN (SELECT TRAN_SEQ_NO,ITEM_SEQ_NO, SUM(QTY*UNIT_DISCOUNT_AMT) AS SALE_TOT_DISC_VAL 
					FROM  SA_TRAN_DISC GROUP BY TRAN_SEQ_NO,ITEM_SEQ_NO )  SD ON SH.TRAN_SEQ_NO = SD.TRAN_SEQ_NO AND SL.ITEM_SEQ_NO =SD.ITEM_SEQ_NO
						GROUP BY SH.TRAN_DATETIME,SH.TRAN_SEQ_NO,SH.TRAN_TYPE,SH.SUB_TRAN_TYPE,SH.STATUS, SH.STORE,SH.TRAN_NO,SH.CASHIER,SH.UPDATE_DATETIME) F
							GROUP BY STORE_CODE)ARC
ON RMS.STORE_CODE = ARC.STORE_CODE
ORDER BY RMS.STORE_CODE
	};
	my $query_handle = $dbh->prepare($query);
	$query_handle->execute();
	
	while (my $y = $query_handle->fetchrow_hashref()) {
	$worksheet->write($a,0, $y->{STORE_CODE},$desc);
	$worksheet->write($a,1, $y->{RMS_SALE_TOT_QTY},$border1);
	$worksheet->write($a,2, $y->{RMS_SALE_NET_VAL},$border1);
	$worksheet->write($a,3, $y->{RMS_SALE_TOT_TAX_VAL},$border1);
	$worksheet->write($a,4, $y->{RMS_SALE_TOT_DISC_VAL},$border1);
	$worksheet->write($a,5, "");
	$worksheet->write($a,6, $y->{ARC_SALE_TOT_QTY},$border1);
	$worksheet->write($a,7, $y->{ARC_SALE_NET_VAL},$border1);
	$worksheet->write($a,8, $y->{ARC_SALE_TOT_TAX_VAL},$border1);
	$worksheet->write($a,9, $y->{ARC_SALE_TOT_DISC_VAL},$border1);
	$worksheet->write($a,10, "");
	$worksheet->write($a,11, $y->{DIFF_SALE_TOT_QTY},$border1);
	$worksheet->write($a,12, $y->{DIFF_SALE_NET_VAL},$border1);
	$worksheet->write($a,13, $y->{DIFF_SALE_TOT_TAX_VAL},$border1);
	$worksheet->write($a,14, $y->{DIFF_SALE_TOT_DISC_VAL},$border1);
	$a++;
}
	
	$query_handle->finish();
}

sub sale_line{
my $worksheet = $workbook->add_worksheet("sale_line");
	$worksheet->set_column( 0, 0, 5 );
	$worksheet->set_column( 1, 8, 12 );
	$worksheet->set_column( 9, 9, 1 );
	$worksheet->set_column( 10, 17, 12 );
	$worksheet->set_column( 18, 18, 1 );
	$worksheet->set_column( 19, 26, 12 );
	$worksheet->set_column( 27, 27, 1 );
	$worksheet->set_column( 28, 35, 12 );
	my $a = 2;
	my $col = 0;
	
	foreach my $i ( "STORE", "RMS_SALE_TOT_QTY", "RMS_SALE_NET_VAL", "RMS_SALE_TOT_TAX_VAL", "RMS_PRODUCT_FULL_PRICE", "RMS_ACTUAL_SELLING_PRICE", "RMS_SALE_TOT_DISC_VAL", "RMS_SALE_MARKDOWN_QTY", "RMS_SALE_MARKDOWN_VAL", 
	"", "ARC_SALE_TOT_QTY", "ARC_SALE_NET_VAL", "ARC_SALE_TOT_TAX_VAL", "ARC_PRODUCT_FULL_PRICE", "ARC_ACTUAL_SELLING_PRICE", "ARC_SALE_TOT_DISC_VAL", "ARC_SALE_MARKDOWN_QTY", "ARC_SALE_MARKDOWN_VAL", 
	"", "DIFF_SALE_TOT_QTY", "DIFF_SALE_NET_VAL", "DIFF_SALE_TOT_TAX_VAL", "DIFF_PRODUCT_FULL_PRICE", "DIFF_ACTUAL_SELLING_PRICE", "DIFF_SALE_TOT_DISC_VAL", "DIFF_SALE_MARKDOWN_QTY", "DIFF_SALE_MARKDOWN_VAL" ) {
		$worksheet->write( 1, $col++, $i, $desc2 );
	}
	
	my $query = qq {
		SELECT RMS.STORE_CODE, RMS_SALE_TOT_QTY, RMS_SALE_NET_VAL, RMS_SALE_TOT_TAX_VAL, RMS_PRODUCT_FULL_PRICE, RMS_ACTUAL_SELLING_PRICE, RMS_SALE_TOT_DISC_VAL, RMS_SALE_MARKDOWN_QTY, RMS_SALE_MARKDOWN_VAL,
ARC.ARC_SALE_TOT_QTY, ARC_SALE_NET_VAL, ARC_SALE_TOT_TAX_VAL, ARC_PRODUCT_FULL_PRICE, ARC_ACTUAL_SELLING_PRICE, ARC_SALE_TOT_DISC_VAL, ARC_SALE_MARKDOWN_QTY, ARC_SALE_MARKDOWN_VAL,
(NVL(RMS_SALE_TOT_QTY,0))-(NVL(ARC_SALE_TOT_QTY,0)) DIFF_SALE_TOT_QTY, 
(NVL(RMS_SALE_NET_VAL,0))-(NVL(ARC_SALE_NET_VAL,0)) DIFF_SALE_NET_VAL, 
(NVL(RMS_SALE_TOT_TAX_VAL,0))-(NVL(ARC_SALE_TOT_TAX_VAL,0)) DIFF_SALE_TOT_TAX_VAL, 
(NVL(RMS_PRODUCT_FULL_PRICE,0))-(NVL(ARC_PRODUCT_FULL_PRICE,0)) DIFF_PRODUCT_FULL_PRICE,
(NVL(RMS_ACTUAL_SELLING_PRICE,0))-(NVL(ARC_ACTUAL_SELLING_PRICE,0)) DIFF_ACTUAL_SELLING_PRICE,
(NVL(RMS_SALE_TOT_DISC_VAL,0))-(NVL(ARC_SALE_TOT_DISC_VAL,0)) DIFF_SALE_TOT_DISC_VAL,
(NVL(RMS_SALE_MARKDOWN_QTY,0))-(NVL(ARC_SALE_MARKDOWN_QTY,0)) DIFF_SALE_MARKDOWN_QTY,
(NVL(RMS_SALE_MARKDOWN_VAL,0))-(NVL(ARC_SALE_MARKDOWN_VAL,0)) DIFF_SALE_MARKDOWN_VAL
FROM
(SELECT SH.STORE AS STORE_CODE,SUM(SL.QTY)AS RMS_SALE_TOT_QTY,SUM(SL.QTY * SL.UNIT_RETAIL)AS RMS_SALE_NET_VAL,
    SUM(TOTAL_IGTAX_AMT)    AS RMS_SALE_TOT_TAX_VAL,SUM(CST.AV_COST)    AS RMS_PRODUCT_FULL_PRICE,SUM(SL.UNIT_RETAIL)    AS RMS_ACTUAL_SELLING_PRICE,
    SUM(SALE_TOT_DISC_VAL)  AS RMS_SALE_TOT_DISC_VAL,SUM( CASE WHEN IC.CLEAR_IND ='Y' THEN NVL(SL.QTY,0)  ELSE 0 END)  AS  RMS_SALE_MARKDOWN_QTY,
    SUM( CASE WHEN IC.CLEAR_IND ='Y' THEN  NVL(SL.QTY,0) * ( NVL(IC.REGULAR_UNIT_RETAIL,0) -NVL( SL.UNIT_RETAIL,0)) ELSE 0 END) AS RMS_SALE_MARKDOWN_VAL
		FROM SA_TRAN_HEAD@RMS SH JOIN SA_TRAN_ITEM  SL ON SH.TRAN_SEQ_NO = SL.TRAN_SEQ_NO AND SH.STORE = SL.STORE
			LEFT JOIN (SELECT TRAN_SEQ_NO,ITEM_SEQ_NO, SUM(QTY*UNIT_DISCOUNT_AMT) AS SALE_TOT_DISC_VAL FROM  SA_TRAN_DISC@RMS GROUP BY TRAN_SEQ_NO,ITEM_SEQ_NO) SD
			ON SH.TRAN_SEQ_NO = SD.TRAN_SEQ_NO AND SL.ITEM_SEQ_NO =SD.ITEM_SEQ_NO
			LEFT JOIN ITEM_LOC_SOH@RMS CST ON SL.ITEM = CST.ITEM AND SL.STORE = CST.LOC LEFT JOIN ITEM_LOC IC ON IC.ITEM = SL.ITEM AND IC.LOC = SL.STORE
			AND IC.CLEAR_IND ='Y' 
		WHERE TRUNC(SH.UPDATE_DATETIME)=(SELECT TO_CHAR(SYSDATE-1,'DD MON YYYY') FROM DUAL) --'02 OCT 2014'
		GROUP BY SH.STORE)RMS
LEFT JOIN
(SELECT SH.STORE AS STORE_CODE,SUM(SL.QTY) AS ARC_SALE_TOT_QTY,SUM(SL.QTY * SL.UNIT_RETAIL) AS ARC_SALE_NET_VAL,
    SUM(TOTAL_IGTAX_AMT)    AS ARC_SALE_TOT_TAX_VAL,SUM(CST.AV_COST)    AS ARC_PRODUCT_FULL_PRICE,SUM(SL.UNIT_RETAIL)    AS ARC_ACTUAL_SELLING_PRICE,
    SUM(SALE_TOT_DISC_VAL)  AS ARC_SALE_TOT_DISC_VAL,SUM( CASE WHEN IC.CLEAR_IND ='Y' THEN NVL(SL.QTY,0)  ELSE 0 END)  AS  ARC_SALE_MARKDOWN_QTY,
    SUM( CASE WHEN IC.CLEAR_IND ='Y' THEN  NVL(SL.QTY,0) * ( NVL(IC.REGULAR_UNIT_RETAIL,0) -NVL( SL.UNIT_RETAIL,0)) ELSE 0 END) AS ARC_SALE_MARKDOWN_VAL
		FROM SA_TRAN_HEAD SH JOIN SA_TRAN_ITEM  SL
			ON SH.TRAN_SEQ_NO = SL.TRAN_SEQ_NO AND SH.STORE = SL.STORE
			LEFT JOIN (SELECT TRAN_SEQ_NO,ITEM_SEQ_NO, SUM(QTY*UNIT_DISCOUNT_AMT) AS SALE_TOT_DISC_VAL 
				FROM  SA_TRAN_DISC GROUP BY TRAN_SEQ_NO,ITEM_SEQ_NO) SD
					ON SH.TRAN_SEQ_NO = SD.TRAN_SEQ_NO AND SL.ITEM_SEQ_NO =SD.ITEM_SEQ_NO
						LEFT JOIN ITEM_LOC_SOH CST ON SL.ITEM = CST.ITEM AND SL.STORE = CST.LOC LEFT JOIN ITEM_LOC IC ON IC.ITEM = SL.ITEM
							AND IC.LOC = SL.STORE AND IC.CLEAR_IND ='Y' GROUP BY SH.STORE)ARC
ON RMS.STORE_CODE = ARC.STORE_CODE							
ORDER BY 1
	};
	my $query_handle = $dbh->prepare($query);
	$query_handle->execute();
	
	while (my $y = $query_handle->fetchrow_hashref()) {
	$worksheet->write($a,0, $y->{STORE_CODE},$desc);
	$worksheet->write($a,1, $y->{RMS_SALE_TOT_QTY},$border1);
	$worksheet->write($a,2, $y->{RMS_SALE_NET_VAL},$border1);
	$worksheet->write($a,3, $y->{RMS_SALE_TOT_TAX_VAL},$border1);
	$worksheet->write($a,4, $y->{RMS_PRODUCT_FULL_PRICE},$border1);
	$worksheet->write($a,5, $y->{RMS_ACTUAL_SELLING_PRICE},$border1);
	$worksheet->write($a,6, $y->{RMS_SALE_TOT_DISC_VAL},$border1);
	$worksheet->write($a,7, $y->{RMS_SALE_MARKDOWN_QTY},$border1);
	$worksheet->write($a,8, $y->{RMS_SALE_MARKDOWN_VAL},$border1);
	$worksheet->write($a,9, "");
	$worksheet->write($a,10, $y->{ARC_SALE_TOT_QTY},$border1);
	$worksheet->write($a,11, $y->{ARC_SALE_NET_VAL},$border1);
	$worksheet->write($a,12, $y->{ARC_SALE_TOT_TAX_VAL},$border1);
	$worksheet->write($a,13, $y->{ARC_PRODUCT_FULL_PRICE},$border1);
	$worksheet->write($a,14, $y->{ARC_ACTUAL_SELLING_PRICE},$border1);
	$worksheet->write($a,15, $y->{ARC_SALE_TOT_DISC_VAL},$border1);
	$worksheet->write($a,16, $y->{ARC_SALE_MARKDOWN_QTY},$border1);
	$worksheet->write($a,17, $y->{ARC_SALE_MARKDOWN_VAL},$border1);
	$worksheet->write($a,18, "");
	$worksheet->write($a,19, $y->{DIFF_SALE_TOT_QTY},$border1);
	$worksheet->write($a,20, $y->{DIFF_SALE_NET_VAL},$border1);
	$worksheet->write($a,21, $y->{DIFF_SALE_TOT_TAX_VAL},$border1);
	$worksheet->write($a,22, $y->{DIFF_PRODUCT_FULL_PRICE},$border1);
	$worksheet->write($a,23, $y->{DIFF_ACTUAL_SELLING_PRICE},$border1);
	$worksheet->write($a,24, $y->{DIFF_SALE_TOT_DISC_VAL},$border1);
	$worksheet->write($a,25, $y->{DIFF_SALE_MARKDOWN_QTY},$border1);
	$worksheet->write($a,26, $y->{DIFF_SALE_MARKDOWN_VAL},$border1);
	$a++;
}
	
	$query_handle->finish();
}

sub sale_invc_tender{
my $worksheet = $workbook->add_worksheet("sale_invc_tender");
	$worksheet->set_column( 0, 0, 5 );
	$worksheet->set_column( 1, 1, 12 );
	$worksheet->set_column( 2, 2, 1 );
	$worksheet->set_column( 3, 3, 12 );
	$worksheet->set_column( 4, 4, 1 );
	$worksheet->set_column( 5, 5, 12 );
	my $a = 2;
	my $col = 0;
	
	foreach my $i ( "STORE", "RMS_TENDER_VAL", "", "ARC_TENDER_VAL", "", "DIFF_TENDER_VAL") {
		$worksheet->write( 1, $col++, $i, $desc2 );
	}
	
	my $query = qq {
	SELECT RMS.STORE_CODE, RMS_TENDER_VAL, ARC_TENDER_VAL, (NVL(RMS_TENDER_VAL,0))-(NVL(ARC_TENDER_VAL,0)) DIFF_TENDER_VAL FROM
	(SELECT ST.STORE AS STORE_CODE, SUM(TENDER_AMT) AS RMS_TENDER_VAL
		FROM SA_TRAN_HEAD@RMS SH
		JOIN SA_TRAN_TENDER@RMS ST
		ON SH.TRAN_SEQ_NO = ST.TRAN_SEQ_NO
	  WHERE TRUNC(SH.UPDATE_DATETIME)=(SELECT TO_CHAR(SYSDATE-1,'DD MON YYYY') FROM DUAL) --'02 OCT 2014'
		GROUP BY ST.STORE)RMS
	LEFT JOIN
	(SELECT ST.STORE AS STORE_CODE, SUM(TENDER_AMT) AS ARC_TENDER_VAL
		FROM SA_TRAN_HEAD SH
		JOIN SA_TRAN_TENDER ST
		ON SH.TRAN_SEQ_NO = ST.TRAN_SEQ_NO
	  GROUP BY ST.STORE)ARC
	ON RMS.STORE_CODE = ARC.STORE_CODE
	ORDER BY 1
	};
	my $query_handle = $dbh->prepare($query);
	$query_handle->execute();
	
	while (my $y = $query_handle->fetchrow_hashref()) {
	$worksheet->write($a,0, $y->{STORE_CODE},$desc);
	$worksheet->write($a,1, $y->{RMS_TENDER_VAL},$border1);
	$worksheet->write($a,2, "");
	$worksheet->write($a,3, $y->{ARC_TENDER_VAL},$border1);
	$worksheet->write($a,4, "");
	$worksheet->write($a,5, $y->{DIFF_TENDER_VAL},$border1);
	$a++;
}
	
	$query_handle->finish();
}

sub sale_invc_discount{
my $worksheet = $workbook->add_worksheet("sale_invc_discount");
	$worksheet->set_column( 0, 0, 5 );
	$worksheet->set_column( 1, 2, 12 );
	$worksheet->set_column( 3, 3, 1 );
	$worksheet->set_column( 4, 5, 12 );
	$worksheet->set_column( 6, 6, 1 );
	$worksheet->set_column( 7, 8, 12 );
	$worksheet->set_column( 9, 9, 1 );
	$worksheet->set_column( 10, 11, 12 );
	my $a = 2;
	my $col = 0;
	
	foreach my $i ( "STORE", "RMS_SALE_TOT_DISC_QTY", "RMS_SALE_TOT_DISC_VAL", "", "ARC_SALE_TOT_DISC_QTY", "ARC_SALE_TOT_DISC_VAL", "", "DIFF_SALE_TOT_DISC_QTY", "DIFF_SALE_TOT_DISC_VAL" ) {
		$worksheet->write( 1, $col++, $i, $desc2 );
	}
	
	my $query = qq {
	SELECT RMS.STORE_CODE, RMS_SALE_TOT_DISC_QTY, RMS_SALE_TOT_DISC_VAL, ARC_SALE_TOT_DISC_QTY, ARC_SALE_TOT_DISC_VAL,
(NVL(RMS_SALE_TOT_DISC_QTY,0))-(NVL(ARC_SALE_TOT_DISC_QTY,0)) DIFF_SALE_TOT_DISC_QTY,
(NVL(RMS_SALE_TOT_DISC_VAL,0))-(NVL(ARC_SALE_TOT_DISC_VAL,0)) DIFF_SALE_TOT_DISC_VAL 
FROM
(SELECT STORE_CODE, SUM(SALE_TOT_DISC_QTY) AS RMS_SALE_TOT_DISC_QTY,	SUM(SALE_TOT_DISC_VAL) AS RMS_SALE_TOT_DISC_VAL
		FROM 
		(SELECT
		    TO_DATE(TO_CHAR(TRAN_DATETIME,'MM/DD/YYYY'),'MM/DD/YYYY')    AS TRANS_DATE,
		    CASE WHEN LENGTH(SH.TRAN_NO) >=10 THEN SUBSTR(SH.TRAN_NO,1,3) ELSE SUBSTR(LPAD(SH.TRAN_NO,10,'0'),1,3) END   AS TILL_CODE,
		    SH.TRAN_SEQ_NO    AS INVOICE_NO,
		    TRAN_TYPE    AS SALE_INVC_TYPE,
		    DISC_TYPE      AS DISCOUNT_TYPE_CODE,
		    SD.STORE    AS STORE_CODE,
		    SD.QTY AS SALE_TOT_DISC_QTY,
		    (SD.QTY*SD.UNIT_DISCOUNT_AMT) AS SALE_TOT_DISC_VAL
		FROM SA_TRAN_HEAD@RMS SH
		JOIN SA_TRAN_DISC@RMS SD
		ON SH.TRAN_SEQ_NO = SD.TRAN_SEQ_NO
		WHERE TRUNC(SH.UPDATE_DATETIME)=(SELECT TO_CHAR(SYSDATE-1,'DD MON YYYY') FROM DUAL) --'02 OCT 2014'
		) SID 
		GROUP BY 
			STORE_CODE)RMS
LEFT JOIN
(SELECT STORE_CODE, SUM(SALE_TOT_DISC_QTY) AS ARC_SALE_TOT_DISC_QTY, SUM(SALE_TOT_DISC_VAL) AS ARC_SALE_TOT_DISC_VAL
		FROM 
		(SELECT
		    TO_DATE(TO_CHAR(TRAN_DATETIME,'MM/DD/YYYY'),'MM/DD/YYYY')    AS TRANS_DATE,
		    CASE WHEN LENGTH(SH.TRAN_NO) >=10 THEN SUBSTR(SH.TRAN_NO,1,3) ELSE SUBSTR(LPAD(SH.TRAN_NO,10,'0'),1,3) END   AS TILL_CODE,
		    SH.TRAN_SEQ_NO    AS INVOICE_NO,
		    TRAN_TYPE    AS SALE_INVC_TYPE,
		    DISC_TYPE      AS DISCOUNT_TYPE_CODE,
		    SD.STORE    AS STORE_CODE,
		    SD.QTY AS SALE_TOT_DISC_QTY,
		    (SD.QTY*SD.UNIT_DISCOUNT_AMT) AS SALE_TOT_DISC_VAL
		FROM SA_TRAN_HEAD@RMS SH
		JOIN SA_TRAN_DISC@RMS SD
		ON SH.TRAN_SEQ_NO = SD.TRAN_SEQ_NO
		) SID 
		GROUP BY 
			STORE_CODE)ARC
ON RMS.STORE_CODE = ARC.STORE_CODE
ORDER BY 1
	};
	my $query_handle = $dbh->prepare($query);
	$query_handle->execute();
	
	while (my $y = $query_handle->fetchrow_hashref()) {
	$worksheet->write($a,0, $y->{STORE_CODE},$desc);
	$worksheet->write($a,1, $y->{RMS_SALE_TOT_DISC_QTY},$border1);
	$worksheet->write($a,2, $y->{RMS_SALE_TOT_DISC_VAL},$border1);
	$worksheet->write($a,3, "");
	$worksheet->write($a,4, $y->{ARC_SALE_TOT_DISC_QTY},$border1);
	$worksheet->write($a,5, $y->{ARC_SALE_TOT_DISC_VAL},$border1);
	$worksheet->write($a,6, "");
	$worksheet->write($a,7, $y->{DIFF_SALE_TOT_DISC_QTY},$border1);
	$worksheet->write($a,8, $y->{DIFF_SALE_TOT_DISC_VAL},$border1);
	$a++;
}
	
	$query_handle->finish();
}

sub sale_line_discount{
my $worksheet = $workbook->add_worksheet("sale_line_discount");
	$worksheet->set_column( 0, 0, 5 );
	$worksheet->set_column( 1, 3, 12 );
	$worksheet->set_column( 4, 4, 1 );
	$worksheet->set_column( 5, 7, 12 );
	$worksheet->set_column( 8, 8, 1 );
	$worksheet->set_column( 9, 11, 12 );
	$worksheet->set_column( 12, 12, 1 );
	$worksheet->set_column( 13, 15, 12 );
	my $a = 2;
	my $col = 0;
	
	foreach my $i ( "STORE", "RMS_SALE_TOT_QTY", "RMS_SALE_TOT_DISC_QTY", "RMS_SALE_TOT_DISC_VAL", "", "ARC_SALE_TOT_QTY", "ARC_SALE_TOT_DISC_QTY", "ARC_SALE_TOT_DISC_VAL", "", "DIFF_SALE_TOT_QTY", "DIFF_SALE_TOT_DISC_QTY", "DIFF_SALE_TOT_DISC_VAL" ) {
		$worksheet->write( 1, $col++, $i, $desc2 );
	}
	
	my $query = qq {
	SELECT RMS.STORE_CODE, RMS_SALE_TOT_QTY, RMS_SALE_TOT_DISC_QTY, RMS_SALE_TOT_DISC_VAL, ARC_SALE_TOT_QTY, ARC_SALE_TOT_DISC_QTY, ARC_SALE_TOT_DISC_VAL, 
	(NVL(RMS_SALE_TOT_QTY,0))-(NVL(ARC_SALE_TOT_QTY,0)) DIFF_SALE_TOT_QTY,
	(NVL(RMS_SALE_TOT_DISC_QTY,0))-(NVL(ARC_SALE_TOT_DISC_QTY,0)) DIFF_SALE_TOT_DISC_QTY,
	(NVL(RMS_SALE_TOT_DISC_VAL,0))-(NVL(ARC_SALE_TOT_DISC_VAL,0)) DIFF_SALE_TOT_DISC_VAL
	FROM
(SELECT STORE_CODE, SUM (SALE_TOT_QTY) AS RMS_SALE_TOT_QTY, SUM (SALE_TOT_DISC_QTY) AS RMS_SALE_TOT_DISC_QTY, SUM (SALE_TOT_DISC_VAL) AS RMS_SALE_TOT_DISC_VAL
FROM 
 (
 SELECT TO_DATE(TO_CHAR (SH.TRAN_DATETIME, 'MM/DD/YYYY'),'MM/DD/YYYY') AS TRANS_DATE,
  CASE WHEN LENGTH(SH.TRAN_NO) >=10 THEN SUBSTR(SH.TRAN_NO,1,3) ELSE SUBSTR(LPAD(SH.TRAN_NO,10,'0'),1,3) END   AS TILL_CODE,
         SD.DISC_TYPE AS DISCOUNT_TYPE_CODE,
         SH.TRAN_SEQ_NO AS INVOICE_NO,
         SD.ITEM_SEQ_NO AS LINE_NO,
         SD.STORE AS STORE_CODE,
         SL.ITEM AS PRODUCT_CODE,
         SL.QTY AS SALE_TOT_QTY,
         SD.QTY AS SALE_TOT_DISC_QTY,
         (SD.QTY * SD.UNIT_DISCOUNT_AMT) AS SALE_TOT_DISC_VAL,
         CASE WHEN SD.PROMOTION IS NULL THEN 1 ELSE 0 END AS PROMO_FLG
    FROM SA_TRAN_HEAD@RMS SH
         JOIN SA_TRAN_ITEM@RMS SL
            ON SH.TRAN_SEQ_NO = SL.TRAN_SEQ_NO
         JOIN SA_TRAN_DISC@RMS SD
            ON SH.TRAN_SEQ_NO = SD.TRAN_SEQ_NO AND SL.ITEM_SEQ_NO = SD.ITEM_SEQ_NO
			WHERE TRUNC(SH.UPDATE_DATETIME)=(SELECT TO_CHAR(SYSDATE-1,'DD MON YYYY') FROM DUAL) --'02 OCT 2014'
             ) SLD
GROUP BY STORE_CODE)RMS
LEFT JOIN
(SELECT STORE_CODE, SUM (SALE_TOT_QTY) AS ARC_SALE_TOT_QTY, SUM (SALE_TOT_DISC_QTY) AS ARC_SALE_TOT_DISC_QTY, SUM (SALE_TOT_DISC_VAL) AS ARC_SALE_TOT_DISC_VAL
FROM 
 (SELECT TO_DATE(TO_CHAR (SH.TRAN_DATETIME, 'MM/DD/YYYY'),'MM/DD/YYYY') AS TRANS_DATE,
  CASE WHEN LENGTH(SH.TRAN_NO) >=10 THEN SUBSTR(SH.TRAN_NO,1,3) ELSE SUBSTR(LPAD(SH.TRAN_NO,10,'0'),1,3) END   AS TILL_CODE,
         SD.DISC_TYPE AS DISCOUNT_TYPE_CODE,
         SH.TRAN_SEQ_NO AS INVOICE_NO,
         SD.ITEM_SEQ_NO AS LINE_NO,
         SD.STORE AS STORE_CODE,
         SL.ITEM AS PRODUCT_CODE,
         SL.QTY AS SALE_TOT_QTY,
         SD.QTY AS SALE_TOT_DISC_QTY,
         (SD.QTY * SD.UNIT_DISCOUNT_AMT) AS SALE_TOT_DISC_VAL,
         CASE WHEN SD.PROMOTION IS NULL THEN 1 ELSE 0 END AS PROMO_FLG
    FROM SA_TRAN_HEAD SH
         JOIN SA_TRAN_ITEM SL
            ON SH.TRAN_SEQ_NO = SL.TRAN_SEQ_NO
         JOIN SA_TRAN_DISC SD
            ON SH.TRAN_SEQ_NO = SD.TRAN_SEQ_NO AND SL.ITEM_SEQ_NO = SD.ITEM_SEQ_NO
             ) SLD GROUP BY STORE_CODE)ARC
ON RMS.STORE_CODE = ARC.STORE_CODE
ORDER BY 1	
	};
	my $query_handle = $dbh->prepare($query);
	$query_handle->execute();
	
	while (my $y = $query_handle->fetchrow_hashref()) {
	$worksheet->write($a,0, $y->{STORE_CODE},$desc);
	$worksheet->write($a,1, $y->{RMS_SALE_TOT_QTY},$border1);
	$worksheet->write($a,2, $y->{RMS_SALE_TOT_DISC_QTY},$border1);
	$worksheet->write($a,3, $y->{RMS_SALE_TOT_DISC_VAL},$border1);
	$worksheet->write($a,4, "");
	$worksheet->write($a,5, $y->{ARC_SALE_TOT_QTY},$border1);
	$worksheet->write($a,6, $y->{ARC_SALE_TOT_DISC_QTY},$border1);
	$worksheet->write($a,7, $y->{ARC_SALE_TOT_DISC_VAL},$border1);
	$worksheet->write($a,8, "");
	$worksheet->write($a,9, $y->{DIFF_SALE_TOT_QTY},$border1);
	$worksheet->write($a,10, $y->{DIFF_SALE_TOT_DISC_QTY},$border1);
	$worksheet->write($a,11, $y->{DIFF_SALE_TOT_DISC_VAL},$border1);
	$a++;
}
	
	$query_handle->finish();
}


sub mail {

my $cc;
my $bcc;
GetOptions( 'cc=s' => \$cc, 'bcc=s' => \$bcc, );

my( $to, $subject, $msgbody_file, $attachment_file ) = @ARGV;

$to = 'kent.mamalias@metrogaisano.com, lea.gonzaga@metrogaisano.com, annalyn.conde@metrogaisano.com, cham.burgos@metrogaisano.com, rex.cabanilla@metrogaisano.com';

$from = 'Report Mailer<report.mailer@metrogaisano.com>';
		
$subject = 'ARC BI VALIDATION RMS TO ARC - DATA AS OF ' . $as_of;

$msgbody_file = 'message_BI.txt';

$attachment_file = "ARC BI VALIDATION RMS TO ARC - Summary (as of $as_of).xlsx";

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





