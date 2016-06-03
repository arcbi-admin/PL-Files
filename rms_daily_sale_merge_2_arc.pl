use DBI;
use DBD::Oracle qw(:ora_types);
use Text::CSV_XS;

# &rms_daily_sale_for_bi;
# &rms_daily_sale;

sub rms_daily_sale_for_bi {

my $hostname = "10.128.0.42";
my $sid = "MGRMSP";
my $port = '1521';
my $uid = 'rmsprd';
my $pw = 'noida123';

# my $hostname = "10.128.4.23";
# my $sid = "MGRMST";
# my $port = '1521';
# my $uid = 'rmsprd';
# my $pw = 'vicsal123';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw, { RaiseError => 1, AutoCommit => 0 }) or die "Unable to connect: $DBI::errstr";

my $csv = Text::CSV_XS->new ({ binary => 1, eol => $/ });
 open my $fh, ">", "rms_daily_sale.csv" or die "rms_daily_sale.csv: $!";

$test = qq{ 
SELECT SALE.TRANS_DATE, SALE.STORE STORE_CODE, SALE.GROUP_NO DEPARTMENT_CODE, SALE.DEPT SUBDEPT, SUM(SALE.SALE_QTY) SALE_QTY, SUM(SALE.ACTUAL_AMT) SALE_AMT_NO_DEDUCTIONS, SUM(SALE.TAX_AMT) TAX_AMT, SUM(NVL(SALE.DISC_AMT,0)) DISC_AMT, SUM(SALE.ACTUAL_AMT-(NVL(SALE.DISC_AMT,0))-(NVL(SALE.TAX_AMT,0))) NET_SALE_WO_VAT_DISC FROM
      (SELECT SALE.TRANS_DATE, SALE.STORE, SALE.GROUP_NO, SALE.GROUP_NAME, SALE.DEPT, SALE.ITEM_SEQ_NO, SALE.ITEM, SALE.IGTAX_RATE, SALE.SALE_QTY, SALE.ACTUAL_AMT, SALE.TAX_AMT, SALE.TAX_TAX, DISC.DISC_AMT, 
      ((nvl(SALE.ACTUAL_AMT,0))-(nvl(DISC.DISC_AMT,0))) AS TST,
      ((nvl(SALE.ACTUAL_AMT,0))-((nvl(SALE.ACTUAL_AMT,0))*((nvl(SALE.IGTAX_RATE,0))/(100 + (nvl(SALE.IGTAX_RATE,0))))))-((nvl(DISC.DISC_AMT,0))-((nvl(DISC.DISC_AMT,0))*((nvl(SALE.IGTAX_RATE,0))/(100 + (nvl(SALE.IGTAX_RATE,0)))))) AS TST2 	
      FROM
          (SELECT TRUNC(H.TRAN_DATETIME) TRANS_DATE, H.STORE, H.TRAN_SEQ_NO, H.TRAN_NO, G.GROUP_NO, G.GROUP_NAME, I.DEPT, I.ITEM_SEQ_NO, I.ITEM, CASE WHEN I.REF_NO5 IN(0707, 0709) THEN 0 ELSE TAX.IGTAX_RATE END AS IGTAX_RATE, SUM(I.QTY) SALE_QTY, SUM(I.QTY*I.UNIT_RETAIL) ACTUAL_AMT, SUM(I.TOTAL_IGTAX_AMT) TAX_AMT, SUM(TAX.TOTAL_IGTAX_AMT) TAX_TAX
          FROM SA_TRAN_HEAD H 
            JOIN SA_TRAN_ITEM I ON H.STORE = I.STORE AND H.TRAN_SEQ_NO = I.TRAN_SEQ_NO
            LEFT JOIN SA_TRAN_IGTAX TAX ON H.STORE = TAX.STORE AND H.TRAN_SEQ_NO = TAX.TRAN_SEQ_NO AND I.ITEM_SEQ_NO=TAX.ITEM_SEQ_NO
            JOIN DEPS ON I.DEPT = DEPS.DEPT
            JOIN GROUPS G ON DEPS.GROUP_NO = G.GROUP_NO
          WHERE TRUNC(TRAN_DATETIME) = '04-NOV-14' --(SELECT TO_CHAR(SYSDATE-1,'DD-MON-YY') FROM DUAL)
            AND H.SUB_TRAN_TYPE IN ('SALE','LAYCMP','RETURN') AND H.STATUS = 'P' AND I.ITEM IS NOT NULL
          GROUP BY TRAN_DATETIME, H.STORE, H.TRAN_SEQ_NO, H.TRAN_NO, G.GROUP_NO, G.GROUP_NAME, I.DEPT, I.ITEM_SEQ_NO, I.ITEM, case when i.ref_no5 in(0707, 0709) then 0 else TAX.IGTAX_RATE end)SALE
        LEFT JOIN
          (SELECT STORE, TRAN_SEQ_NO, ITEM_SEQ_NO, SUM(QTY*UNIT_DISCOUNT_AMT) DISC_AMT FROM SA_TRAN_DISC GROUP BY STORE, TRAN_SEQ_NO, ITEM_SEQ_NO)DISC 
        ON SALE.STORE = DISC.STORE AND SALE.TRAN_SEQ_NO = DISC.TRAN_SEQ_NO AND SALE.ITEM_SEQ_NO = DISC.ITEM_SEQ_NO
      )SALE
    GROUP BY SALE.TRANS_DATE, SALE.STORE, SALE.GROUP_NO, SALE.DEPT
};

my $sth = $dbh->prepare ($test);
 $sth->execute;
 $csv->print ($fh, $sth->{NAME_uc});
 while (my $row = $sth->fetch) {
     $csv->print ($fh, $row) or $csv->error_diag;
     }
 close $fh or die "rms_daily_sale.csv: $!";

$dbh->disconnect;

}

sub rms_daily_sale {

my $hostname = "10.128.0.220";
my $sid = "METROBIP";
my $port = '1521';
my $uid = 'ARCMA';
my $pw = 'arcma';

my $dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw, { RaiseError => 1, AutoCommit => 0 }) or die "Unable to connect: $DBI::errstr";

my $sth_upsert = $dbh->prepare( q{
MERGE INTO METRO_IT_RMS_DAILY_SALE2 USING dual ON ( TRAN_DATE = ? AND STORE_CODE = ? AND DEPARTMENT_CODE = ? AND SUBDEPARTMENT_CODE = ? )
WHEN MATCHED THEN 
UPDATE SET SALE_QTY = ?, SALE_AMT_NO_DEDUCTIONS = ?, TAX_AMT = ?, DISC_AMT = ?, NET_SALE_WO_VAT_DISC = ?
WHEN NOT MATCHED THEN 
INSERT (STORE_CODE, DEPARTMENT_CODE, SUBDEPARTMENT_CODE, SALE_QTY, SALE_AMT_NO_DEDUCTIONS, TAX_AMT, DISC_AMT, NET_SALE_WO_VAT_DISC, TRAN_DATE) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ? )
});

open FH2, "<rms_daily_sale.csv" or die "Unable to open rms_daily_sale.csv: $!";
while (<FH2>) {
	chomp;
    my ( $tran_date, $store_code, $department_code, $subdepartment_code, $sale_qty, $sale_amt_no_deductions, $tax_amt, $disc_amt, $net_sale_wo_vat_disc ) = split /,/;
	
	$sth_upsert->execute( $tran_date, $store_code, $department_code, $subdepartment_code, $sale_qty, $sale_amt_no_deductions, $tax_amt, $disc_amt, $net_sale_wo_vat_disc, $store_code, $department_code, $subdepartment_code, $sale_qty, $sale_amt_no_deductions, $tax_amt, $disc_amt, $net_sale_wo_vat_disc, $tran_date );
}

close FH2;

print "Done with Upsert... \n";

$dbh->commit;
$dbh->disconnect;

}







