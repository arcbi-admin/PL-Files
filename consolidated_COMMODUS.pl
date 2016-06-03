use DBI;
use DBD::Oracle qw(:ora_types);
use Text::CSV_XS;
use DBConnector;
#use DBConnectorCommodus;


&iupc;
# &sls_inventory;
# &sls_inventory_repnonrep;
# &sls_inventory_bystore;
# &sls_inventory_bysku;


$dbh->commit;
$dbh->disconnect;

sub iupc {

####################### 	 SLS + INVENTORY 	#######################

#my $create_table = $dbhCommodus->prepare( q{
# my $create_table = $dbh->prepare( q{
# CREATE TABLE ARC_DW_MA.TEMP_MM_KENT2 (
# DEPARTMENT_CODE CHAR(15) DEFAULT 0, 
# DEPARTMENT_DESC VARCHAR2(100), 
# CLASS_CODE CHAR(15) DEFAULT 0, 
# CLASS_DESC VARCHAR2(100), 
# SUBCLASS_CODE CHAR(15) DEFAULT 0, 
# SUBCLASS_DESC VARCHAR2(100), 
# PROD_TYPE_CODE NUMBER(15) NOT NULL, 
# PROD_TYPE_DESC VARCHAR2(100), 
# SLSTY_YTD NUMBER(20) DEFAULT 0, 
# SLSLY_YTD NUMBER(20) DEFAULT 0 , 
# SLS_GROWTH_YTD NUMBER(20) DEFAULT 0, 
# MARGINTY_YTD NUMBER(20) DEFAULT 0, 
# MARGINLY_YTD NUMBER(20) DEFAULT 0, 
# MRGN_GROWTH_YTD NUMBER(20) DEFAULT 0, 
# GM_TY NUMBER(20) DEFAULT 0, 
# GM_LY NUMBER(20) DEFAULT 0, 
# SLSTY_MTD NUMBER(20) DEFAULT 0, 
# SLSLY_MTD NUMBER(20) DEFAULT 0 , 
# SLS_GROWTH_MTD NUMBER(20) DEFAULT 0, 
# MARGINTY_MTD NUMBER(20) DEFAULT 0, 
# MARGINLY_MTD NUMBER(20) DEFAULT 0, 
# MRGN_GROWTH_MTD NUMBER(20) DEFAULT 0, 
# SLSTY_WTD NUMBER(20) DEFAULT 0, 
# SLSLY_WTD NUMBER(20) DEFAULT 0 , 
# SLS_GROWTH_WTD NUMBER(20) DEFAULT 0, 
# MARGINTY_WTD NUMBER(20) DEFAULT 0, 
# MARGINLY_WTD NUMBER(20) DEFAULT 0, 
# MRGN_GROWTH_WTD NUMBER(20) DEFAULT 0, 
# ISCLAS CHAR(15) DEFAULT 0, 
# INVCOST NUMBER(20) DEFAULT 0, 
# INVRETL NUMBER(20) DEFAULT 0, 
# BMVAL NUMBER(20) DEFAULT 0, 
# YRECEIPT NUMBER(20) DEFAULT 0, 
# MD NUMBER(20) DEFAULT 0, UNIQUE(PROD_TYPE_CODE)
# ) TABLESPACE USERS
# });
# $create_table->execute();

my $truncate = $dbh->prepare( qq{ 
TRUNCATE TABLE TEMP_IUPC 
});
$truncate->execute();

print "Done truncating TEMP_IUPC... \nPreparing to Insert... \n";

my $sth_insert = $dbh->prepare( q{
INSERT INTO TEMP_IUPC (IUPC, INUMBR)
VALUES ( ?, ? ) 
});
  
open FH1, "<iupc.csv" or die "Unable to open iupc.csv: $!";
while (<FH1>) {
	chomp;
    my ( $iupc, $inumbr ) = split (/,/);
	
	$sth_insert->execute( $iupc, $inumbr );	
}
close FH1;

$dbh->commit;

print "Done with Insert...";

}


sub sls_inventory {

####################### 	 SLS + INVENTORY 	#######################

#my $create_table = $dbhCommodus->prepare( q{
# my $create_table = $dbh->prepare( q{
# CREATE TABLE ARC_DW_MA.TEMP_MM_KENT2 (
# DEPARTMENT_CODE CHAR(15) DEFAULT 0, 
# DEPARTMENT_DESC VARCHAR2(100), 
# CLASS_CODE CHAR(15) DEFAULT 0, 
# CLASS_DESC VARCHAR2(100), 
# SUBCLASS_CODE CHAR(15) DEFAULT 0, 
# SUBCLASS_DESC VARCHAR2(100), 
# PROD_TYPE_CODE NUMBER(15) NOT NULL, 
# PROD_TYPE_DESC VARCHAR2(100), 
# SLSTY_YTD NUMBER(20) DEFAULT 0, 
# SLSLY_YTD NUMBER(20) DEFAULT 0 , 
# SLS_GROWTH_YTD NUMBER(20) DEFAULT 0, 
# MARGINTY_YTD NUMBER(20) DEFAULT 0, 
# MARGINLY_YTD NUMBER(20) DEFAULT 0, 
# MRGN_GROWTH_YTD NUMBER(20) DEFAULT 0, 
# GM_TY NUMBER(20) DEFAULT 0, 
# GM_LY NUMBER(20) DEFAULT 0, 
# SLSTY_MTD NUMBER(20) DEFAULT 0, 
# SLSLY_MTD NUMBER(20) DEFAULT 0 , 
# SLS_GROWTH_MTD NUMBER(20) DEFAULT 0, 
# MARGINTY_MTD NUMBER(20) DEFAULT 0, 
# MARGINLY_MTD NUMBER(20) DEFAULT 0, 
# MRGN_GROWTH_MTD NUMBER(20) DEFAULT 0, 
# SLSTY_WTD NUMBER(20) DEFAULT 0, 
# SLSLY_WTD NUMBER(20) DEFAULT 0 , 
# SLS_GROWTH_WTD NUMBER(20) DEFAULT 0, 
# MARGINTY_WTD NUMBER(20) DEFAULT 0, 
# MARGINLY_WTD NUMBER(20) DEFAULT 0, 
# MRGN_GROWTH_WTD NUMBER(20) DEFAULT 0, 
# ISCLAS CHAR(15) DEFAULT 0, 
# INVCOST NUMBER(20) DEFAULT 0, 
# INVRETL NUMBER(20) DEFAULT 0, 
# BMVAL NUMBER(20) DEFAULT 0, 
# YRECEIPT NUMBER(20) DEFAULT 0, 
# MD NUMBER(20) DEFAULT 0, UNIQUE(PROD_TYPE_CODE)
# ) TABLESPACE USERS
# });
# $create_table->execute();

#my $truncate = $dbhCommodus->prepare( qq{ 
my $truncate = $dbh->prepare( qq{ 
TRUNCATE TABLE TEMP_MM_KENT2 
});
$truncate->execute();

print "Done truncating TEMP_MM_KENT2... \nPreparing to Insert... \n";

#my $sth_insert = $dbhCommodus->prepare( q{
my $sth_insert = $dbh->prepare( q{
INSERT INTO TEMP_MM_KENT2 (DEPARTMENT_CODE, DEPARTMENT_DESC, CLASS_CODE, CLASS_DESC, SUBCLASS_CODE, SUBCLASS_DESC, PROD_TYPE_CODE, PROD_TYPE_DESC, SLSTY_YTD, SLSLY_YTD, SLS_GROWTH_YTD, MARGINTY_YTD, MARGINLY_YTD, MRGN_GROWTH_YTD, GM_TY, GM_LY, SLSTY_MTD, SLSLY_MTD, SLS_GROWTH_MTD, MARGINTY_MTD, MARGINLY_MTD, MRGN_GROWTH_MTD, SLSTY_WTD, SLSLY_WTD, SLS_GROWTH_WTD, MARGINTY_WTD, MARGINLY_WTD, MRGN_GROWTH_WTD, ISCLAS, INVCOST, INVRETL, BMVAL, YRECEIPT, MD)
VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? ) 
});
  
open FH1, "<mm_all_sls.csv" or die "Unable to open mm_all_sls.csv: $!";
while (<FH1>) {
	chomp;
    my ( $department_code, $department_desc, $class_code, $class_desc, $subclass_code, $subclass_desc, $prod_type_code, $prod_type_desc, $slsty_ytd, $slsly_ytd, $sls_growth_ytd, $marginty_ytd, $marginly_ytd, $mrgn_growth_ytd, $gm_ty, $gm_ly, $slsty_mtd, $slsly_mtd, $sls_growth_mtd, $marginty_mtd, $marginly_mtd, $mrgn_growth_mtd, $slsty_wtd, $slsly_wtd, $sls_growth_wtd, $marginty_wtd, $marginly_wtd, $mrgn_growth_wtd, $isclas, $invcost, $invretl, $bmval, $yreceipt, $md ) = split (/,/);
	
	$sth_insert->execute( $department_code, $department_desc, $class_code, $class_desc, $subclass_code, $subclass_desc, $prod_type_code, $prod_type_desc, $slsty_ytd, $slsly_ytd, $sls_growth_ytd, $marginty_ytd, $marginly_ytd, $mrgn_growth_ytd, $gm_ty, $gm_ly, $slsty_mtd, $slsly_mtd, $sls_growth_mtd, $marginty_mtd, $marginly_mtd, $mrgn_growth_mtd, $slsty_wtd, $slsly_wtd, $sls_growth_wtd, $marginty_wtd, $marginly_wtd, $mrgn_growth_wtd, $isclas, $invcost, $invretl, $bmval, $yreceipt, $md );	
}
close FH1;

#$dbhCommodus->commit;
$dbh->commit;

print "Done with Insert... \nPreparing to Upsert... \n";

#my $sth_upsert = $dbhCommodus->prepare( q{
my $sth_upsert = $dbh->prepare( q{
MERGE INTO TEMP_MM_KENT2 USING dual ON ( PROD_TYPE_CODE= ? )
WHEN MATCHED THEN 
UPDATE SET ISCLAS = ?, INVCOST = ?, INVRETL = ?, BMVAL = ?, YRECEIPT = ?, MD = ?
WHEN NOT MATCHED THEN 
INSERT (DEPARTMENT_CODE, CLASS_CODE, SUBCLASS_CODE, PROD_TYPE_CODE, ISCLAS, INVCOST, INVRETL, BMVAL, YRECEIPT, MD) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ? )
});

open FH2, "<mm_all_inventory_and_receipt.csv" or die "Unable to open mm_all_inventory_and_receipt.csv: $!";
while (<FH2>) {
	chomp;
    my ( $idept, $isdept, $iclas, $isclas, $invcost, $invretl, $bmval, $yreceipt, $md ) = split /,/;
	
	$sth_upsert->execute( $isclas, $isclas, $invcost, $invretl, $bmval, $yreceipt, $md, $idept, $isdept, $iclas, $isclas, $isclas, $invcost, $invretl, $bmval, $yreceipt, $md );
}
close FH2;

print "Done with Upsert... \n";

#$dbhCommodus->commit;
$dbh->commit;

}

sub sls_inventory_repnonrep {

###################### 	SLS + INVENTORY , REP/NONREP 	#######################

#my $create_table = $dbhCommodus->prepare( q{
# my $create_table = $dbh->prepare( q{
# CREATE TABLE ARC_DW_MA.TEMP_MM_REPNON_KENT2 (
# DEPARTMENT_CODE CHAR(15) DEFAULT 0, 
# DEPARTMENT_DESC VARCHAR2(100), 
# CLASS_CODE CHAR(15) DEFAULT 0, 
# CLASS_DESC VARCHAR2(100), 
# SUBCLASS_CODE CHAR(15) DEFAULT 0, 
# SUBCLASS_DESC VARCHAR2(100), 
# PROD_TYPE_CODE NUMBER(15) NOT NULL, 
# PROD_TYPE_DESC VARCHAR2(100), 
# SLSTY NUMBER(20) DEFAULT 0, 
# REP_SLSTY NUMBER(20) DEFAULT 0, 
# NONREP_SLSTY NUMBER(20) DEFAULT 0, 
# MARGINTY NUMBER(20) DEFAULT 0, 
# COMP_INVRETL_REP NUMBER(20) DEFAULT 0, 
# COMP_INVCOST_REP NUMBER(20) DEFAULT 0, 
# COMP_INVRETL_NONREP NUMBER(20) DEFAULT 0, 
# COMP_INVCOST_NONREP NUMBER(20) DEFAULT 0, 
# ALL_INVRETL_REP NUMBER(20) DEFAULT 0, 
# ALL_INVCOST_REP NUMBER(20) DEFAULT 0, 
# ALL_INVRETL_NONREP NUMBER(20) DEFAULT 0, 
# ALL_INVCOST_NONREP NUMBER(20) DEFAULT 0, 
# PEND_TOTRETL_1STHALF NUMBER(20) DEFAULT 0, 
# PEND_TOTRETL_2NDHALF NUMBER(20) DEFAULT 0, 
# BMVAL NUMBER(20) DEFAULT 0, UNIQUE(PROD_TYPE_CODE)
# )  TABLESPACE USERS
# });
# $create_table->execute();

#my $truncate = $dbhCommodus->prepare( qq{ 
my $truncate = $dbh->prepare( qq{ 
TRUNCATE TABLE TEMP_MM_REPNON_KENT2 
});
$truncate->execute();

print "Done truncating TEMP_MM_REPNON_KENT2... \nPreparing to Insert... \n";

#my $sth_insert = $dbhCommodus->prepare( q{
my $sth_insert = $dbh->prepare( q{
INSERT INTO TEMP_MM_REPNON_KENT2 (DEPARTMENT_CODE, DEPARTMENT_DESC, CLASS_CODE, CLASS_DESC, SUBCLASS_CODE, SUBCLASS_DESC, PROD_TYPE_CODE, PROD_TYPE_DESC, SLSTY, REP_SLSTY, NONREP_SLSTY, MARGINTY, COMP_INVRETL_REP, COMP_INVCOST_REP, COMP_INVRETL_NONREP, COMP_INVCOST_NONREP, ALL_INVRETL_REP, ALL_INVCOST_REP, ALL_INVRETL_NONREP, ALL_INVCOST_NONREP, PEND_TOTRETL_1STHALF, PEND_TOTRETL_2NDHALF, BMVAL)
VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? ) 
});
  
open FH3, "<mm_sls_repnon.csv" or die "Unable to open mm_sls_repnon.csv: $!";
while (<FH3>) {
	chomp;
    my ( $department_code, $department_desc, $class_code, $class_desc, $subclass_code, $subclass_desc, $prod_type_code, $prod_type_desc, $slsty, $rep_slsty, $nonrep_slsty, $marginty, $comp_invretl_rep, $comp_invcost_rep, $comp_invretl_nonrep, $comp_invcost_nonrep, $all_invretl_rep, $all_invcost_rep, $all_invretl_nonrep, $all_invcost_nonrep, $pend_totretl_1sthalf, $pend_totretl_2ndhalf, $bmval ) = split (/,/);
	
	$sth_insert->execute( $department_code, $department_desc, $class_code, $class_desc, $subclass_code, $subclass_desc, $prod_type_code, $prod_type_desc, $slsty, $rep_slsty, $nonrep_slsty, $marginty, $comp_invretl_rep, $comp_invcost_rep, $comp_invretl_nonrep, $comp_invcost_nonrep, $all_invretl_rep, $all_invcost_rep, $all_invretl_nonrep, $all_invcost_nonrep, $pend_totretl_1sthalf, $pend_totretl_2ndhalf, $bmval );	
}
close FH3;

#$dbhCommodus->commit;
$dbh->commit;

print "Done with Insert... \nPreparing to Upsert... \n";

#my $sth_upsert = $dbhCommodus->prepare( q{
my $sth_upsert = $dbh->prepare( q{
MERGE INTO TEMP_MM_REPNON_KENT2 USING dual ON ( PROD_TYPE_CODE= ? )
WHEN MATCHED THEN 
UPDATE SET  COMP_INVRETL_REP = ?, COMP_INVCOST_REP = ?, COMP_INVRETL_NONREP = ?, COMP_INVCOST_NONREP = ?, ALL_INVRETL_REP = ?, ALL_INVCOST_REP = ?, ALL_INVRETL_NONREP = ?, ALL_INVCOST_NONREP = ?, PEND_TOTRETL_1STHALF = ?, PEND_TOTRETL_2NDHALF = ?, BMVAL = ?
WHEN NOT MATCHED THEN 
INSERT (DEPARTMENT_CODE, CLASS_CODE, SUBCLASS_CODE, PROD_TYPE_CODE, COMP_INVRETL_REP, COMP_INVCOST_REP, COMP_INVRETL_NONREP, COMP_INVCOST_NONREP, ALL_INVRETL_REP, ALL_INVCOST_REP, ALL_INVRETL_NONREP, ALL_INVCOST_NONREP, PEND_TOTRETL_1STHALF, PEND_TOTRETL_2NDHALF, BMVAL) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? )
});

open FH4, "<mm_invlevel_repnon.csv" or die "Unable to open mm_invlevel_repnon.csv: $!";
while (<FH4>) {
	chomp;
    my ( $idept, $isdept, $iclas, $isclas, $comp_invretl_rep, $comp_invcost_rep, $comp_invretl_nonrep, $comp_invcost_nonrep, $all_invretl_rep, $all_invcost_rep, $all_invretl_nonrep, $all_invcost_nonrep, $pend_totretl_1sthalf, $pend_totretl_2ndhalf, $bmval ) = split /,/;
	
	$sth_upsert->execute( $isclas, $comp_invretl_rep, $comp_invcost_rep, $comp_invretl_nonrep, $comp_invcost_nonrep, $all_invretl_rep, $all_invcost_rep, $all_invretl_nonrep, $all_invcost_nonrep, $pend_totretl_1sthalf, $pend_totretl_2ndhalf, $bmval, $idept, $isdept, $iclas, $isclas, $comp_invretl_rep, $comp_invcost_rep, $comp_invretl_nonrep, $comp_invcost_nonrep, $all_invretl_rep, $all_invcost_rep, $all_invretl_nonrep, $all_invcost_nonrep, $pend_totretl_1sthalf, $pend_totretl_2ndhalf, $bmval );
}
close FH4;

print "Done with Upsert... \n";

#$dbhCommodus->commit;
$dbh->commit;

}

sub sls_inventory_bystore {

####################### 	SLS + INVENTORY, BY STORE	#######################

#my $create_table = $dbhCommodus->prepare( q{
# my $create_table = $dbh->prepare( q{
# CREATE TABLE ARC_DW_MA.TEMP_MM_BYSTORE_KENT2(
# STORE_POSITION NUMBER(15) DEFAULT 0, 
# DEPARTMENT_CODE CHAR(15) DEFAULT 0, 
# DEPARTMENT_DESC VARCHAR2(100), 
# CLASS_CODE CHAR(15) DEFAULT 0, 
# CLASS_DESC VARCHAR2(100),
# SUBCLASS_CODE CHAR(15) DEFAULT 0, 
# SUBCLASS_DESC VARCHAR2(100), 
# PROD_TYPE_CODE NUMBER(15) NOT NULL, 
# PROD_TYPE_DESC VARCHAR2(100), 
# SLSTY_YTD NUMBER(20) DEFAULT 0, 
# SLSLY_YTD NUMBER(20) DEFAULT 0, 
# SLS_GROWTH_YTD NUMBER(20) DEFAULT 0, 
# MARGINTY_YTD NUMBER(20) DEFAULT 0, 
# MARGINLY_YTD NUMBER(20) DEFAULT 0, 
# MRGN_GROWTH_YTD NUMBER(20) DEFAULT 0, 
# SLSTY_MTD NUMBER(20) DEFAULT 0, 
# SLSLY_MTD NUMBER(20) DEFAULT 0, 
# SLS_GROWTH_MTD NUMBER(20) DEFAULT 0, 
# MARGINTY_MTD NUMBER(20) DEFAULT 0, 
# MARGINLY_MTD NUMBER(20) DEFAULT 0, 
# MRGN_GROWTH_MTD NUMBER(20) DEFAULT 0, 
# SLSTY_WTD NUMBER(20) DEFAULT 0, 
# SLSLY_WTD NUMBER(20) DEFAULT 0, 
# SLS_GROWTH_WTD NUMBER(20) DEFAULT 0, 
# MARGINTY_WTD NUMBER(20) DEFAULT 0, 
# MARGINLY_WTD NUMBER(20) DEFAULT 0, 
# MRGN_GROWTH_WTD NUMBER(20) DEFAULT 0, 
# INVCOST NUMBER(20) DEFAULT 0, 
# INVRETL NUMBER(20) DEFAULT 0, 
# BMVAL NUMBER(20) DEFAULT 0
# )TABLESPACE USERS
# });
# $create_table->execute();

#my $truncate = $dbhCommodus->prepare( qq{ 
my $truncate = $dbh->prepare( qq{ 
TRUNCATE TABLE TEMP_MM_BYSTORE_KENT2 
});
$truncate->execute();

print "Done truncating TEMP_MM_BYSTORE_KENT2... \nPreparing to Insert... \n";

#my $sth_insert = $dbhCommodus->prepare( q{
my $sth_insert = $dbh->prepare( q{
INSERT INTO TEMP_MM_BYSTORE_KENT2 (STORE_POSITION,DEPARTMENT_CODE, DEPARTMENT_DESC, CLASS_CODE, CLASS_DESC, SUBCLASS_CODE, SUBCLASS_DESC, PROD_TYPE_CODE, PROD_TYPE_DESC, SLSTY_YTD, SLSLY_YTD, SLS_GROWTH_YTD, MARGINTY_YTD, MARGINLY_YTD, MRGN_GROWTH_YTD, SLSTY_MTD, SLSLY_MTD, SLS_GROWTH_MTD, MARGINTY_MTD, MARGINLY_MTD, MRGN_GROWTH_MTD, SLSTY_WTD, SLSLY_WTD, SLS_GROWTH_WTD, MARGINTY_WTD, MARGINLY_WTD, MRGN_GROWTH_WTD, INVCOST, INVRETL, BMVAL)
VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? ) 
});
  
open FH5, "<mm_sls_bystore.csv" or die "Unable to open mm_sls_bystore.csv: $!";
while (<FH5>) {
	chomp;
    my ( $store_position, $department_code, $department_desc, $class_code, $class_desc, $subclass_code, $subclass_desc, $prod_type_code, $prod_type_desc, $slsty_ytd, $slsly_ytd, $sls_growth_ytd, $marginty_ytd, $marginly_ytd, $mrgn_growth_ytd, $slsty_mtd, $slsly_mtd, $sls_growth_mtd, $marginty_mtd, $marginly_mtd, $mrgn_growth_mtd, $slsty_wtd, $slsly_wtd, $sls_growth_wtd, $marginty_wtd, $marginly_wtd, $mrgn_growth_wtd, $invcost, $invretl, $bmval ) = split (/,/);
	
	$sth_insert->execute( $store_position, $department_code, $department_desc, $class_code, $class_desc, $subclass_code, $subclass_desc, $prod_type_code, $prod_type_desc, $slsty_ytd, $slsly_ytd, $sls_growth_ytd, $marginty_ytd, $marginly_ytd, $mrgn_growth_ytd, $slsty_mtd, $slsly_mtd, $sls_growth_mtd, $marginty_mtd, $marginly_mtd, $mrgn_growth_mtd, $slsty_wtd, $slsly_wtd, $sls_growth_wtd, $marginty_wtd, $marginly_wtd, $mrgn_growth_wtd, $invcost, $invretl, $bmval );	
}
close FH5;

#$dbhCommodus->commit;
$dbh->commit;

print "Done with Insert... \nPreparing to Upsert... \n";

#my $sth_upsert = $dbhCommodus->prepare( q{
my $sth_upsert = $dbh->prepare( q{
MERGE INTO TEMP_MM_BYSTORE_KENT2 USING dual ON ( STORE_POSITION = ? AND PROD_TYPE_CODE = ? )
WHEN MATCHED THEN 
UPDATE SET INVCOST = ?, INVRETL = ?, BMVAL = ?
WHEN NOT MATCHED THEN 
INSERT (STORE_POSITION, DEPARTMENT_CODE, CLASS_CODE, SUBCLASS_CODE, PROD_TYPE_CODE, INVCOST, INVRETL, BMVAL) VALUES ( ?, ?, ?, ?, ?, ?, ?, ? )
});

open FH6, "<mm_invlevel_bystore.csv" or die "Unable to open mm_invlevel_bystore.csv: $!";
while (<FH6>) {
	chomp;
    my ( $store_position, $idept, $isdept, $iclas, $isclas, $invcost, $invretl, $bmval ) = split /,/;
	
	$sth_upsert->execute( $store_position, $isclas, $invcost, $invretl, $bmval, $store_position, $idept, $isdept, $iclas, $isclas, $invcost, $invretl, $bmval);
}
close FH6;

print "Done with Upsert... \n";

#$dbhCommodus->commit;
$dbh->commit;

}

sub sls_inventory_bysku {

###################### 	SLS + INVENTORY , SKU 	#######################

#my $create_sku_table = $dbhCommodus->prepare( q{
# my $create_sku_table = $dbh->prepare( q{
# CREATE TABLE ARC_DW_MA.TEMP_MM_SKU_KENT2 (
# DEPARTMENT_CODE CHAR(15) DEFAULT 0, 
# DEPARTMENT_DESC VARCHAR2(100), 
# CLASS_CODE CHAR(15) DEFAULT 0, 
# CLASS_DESC VARCHAR2(100), 
# SUBCLASS_CODE CHAR(15) DEFAULT 0, 
# SUBCLASS_DESC VARCHAR2(100), 
# PROD_TYPE_CODE NUMBER(15) NOT NULL, 
# PROD_TYPE_DESC VARCHAR2(100), 
# SKU_CODE NUMBER(15) NOT NULL, 
# SKU_DESC VARCHAR2(100), 
# SKU_TYPE VARCHAR2(5), 
# SLSTY NUMBER(20) DEFAULT 0, 
# COMP_INVRETL NUMBER(20) DEFAULT 0, 
# COMP_INVCOST NUMBER(20) DEFAULT 0, 
# ALL_INVRETL NUMBER(20) DEFAULT 0, 
# ALL_INVCOST NUMBER(20) DEFAULT 0, 
# BMVAL NUMBER(20) DEFAULT 0, UNIQUE(SKU_CODE)
# ) TABLESPACE USERS
# });
# $create_sku_table->execute();

#my $truncate_sku = $dbhCommodus->prepare( qq{ 
my $truncate_sku = $dbh->prepare( qq{ 
TRUNCATE TABLE TEMP_MM_SKU_KENT2 
});
$truncate_sku->execute();

print "Done truncating TEMP_MM_SKU_KENT2... \nPreparing to Insert... \n";

#my $sth_insert = $dbhCommodus->prepare( q{
my $sth_insert = $dbh->prepare( q{
MERGE INTO TEMP_MM_SKU_KENT2 USING dual ON ( SKU_CODE= ? )
WHEN MATCHED THEN 
UPDATE SET SLSTY = ?
WHEN NOT MATCHED THEN 
INSERT (DEPARTMENT_CODE, DEPARTMENT_DESC, CLASS_CODE, CLASS_DESC, SUBCLASS_CODE, SUBCLASS_DESC, PROD_TYPE_CODE, PROD_TYPE_DESC, SKU_CODE, SKU_DESC, SKU_TYPE, SLSTY, COMP_INVRETL, COMP_INVCOST, ALL_INVRETL, ALL_INVCOST, BMVAL) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? )
});
  
open FH7, "<mm_sls_repnon_sku.csv" or die "Unable to open mm_sls_repnon_sku.csv: $!";
while (<FH7>) {
	chomp;
    my ( $department_code, $department_desc, $class_code, $class_desc, $subclass_code, $subclass_desc, $prod_type_code, $prod_type_desc, $sku_code, $sku_desc, $sku_type, $slsty, $comp_invretl, $comp_invcost, $all_invretl, $all_invcost, $bmval ) = split (/,/);
	
	$sth_insert->execute( $sku_code, $slsty, $department_code, $department_desc, $class_code, $class_desc, $subclass_code, $subclass_desc, $prod_type_code, $prod_type_desc, $sku_code, $sku_desc, $sku_type, $slsty, $comp_invretl, $comp_invcost, $all_invretl, $all_invcost, $bmval );	
}
close FH7;

#$dbhCommodus->commit;
$dbh->commit;

print "Done with Insert... \nPreparing to Upsert... \n";

#my $sth_upsert = $dbhCommodus->prepare( q{ 
my $sth_upsert = $dbh->prepare( q{ 
MERGE INTO TEMP_MM_SKU_KENT2 USING dual ON ( SKU_CODE= ? )
WHEN MATCHED THEN 
UPDATE SET COMP_INVRETL = ?, COMP_INVCOST = ?, ALL_INVRETL = ?, ALL_INVCOST = ?, BMVAL = ?
WHEN NOT MATCHED THEN 
INSERT (DEPARTMENT_CODE, DEPARTMENT_DESC, CLASS_CODE, CLASS_DESC, SUBCLASS_CODE, SUBCLASS_DESC, PROD_TYPE_CODE, PROD_TYPE_DESC, SKU_CODE, SKU_DESC, SKU_TYPE, COMP_INVRETL, COMP_INVCOST, ALL_INVRETL, ALL_INVCOST, BMVAL) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? )
});

open FH8, "<mm_invlevel_repnon_sku.csv" or die "Unable to open mm_invlevel_repnon_sku.csv: $!";
while (<FH8>) {
	chomp;
    my ($idept,$dptnam,$isdept,$sdptnam,$iclas,$clsnam,$isclas,$sclsnam,$sku_code,$idescr,$itmtyp,$comp_invretl,$comp_invcost,$all_invretl,$all_invcost,$bmval ) = split /,/;
	
	$sth_upsert->execute($sku_code,$comp_invretl,$comp_invcost,$all_invretl,$all_invcost,$bmval,$idept,$dptnam,$isdept,$sdptnam,$iclas,$clsnam,$isclas,$sclsnam,$sku_code,$idescr,$itmtyp,$comp_invretl,$comp_invcost,$all_invretl,$all_invcost,$bmval );
}
close FH8;

print "Done with Upsert... \n";
  
#$dbhCommodus->commit;
$dbh->commit;

}




