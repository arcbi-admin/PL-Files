use DBI;
use DBD::Oracle qw(:ora_types);
use Date::Calc qw( Today Day_of_Week Month_to_Text Week_Number );
use DBConnector;


$date = qq{ 
SELECT WEEK_NUMBER_THIS_YEAR, DATE_KEY, TO_CHAR(DATE_FLD, 'DD Mon YYYY') DATE_FLD, WEEK_ST_DATE_KEY, WEEK_END_DATE_KEY 
FROM DIM_DATE 
WHERE DATE_FLD = (SELECT TO_DATE(VALUE,'YYYY-MM-DD') FROM ADMIN_ETL_SUMMARY)
 };

my $sth_date_1 = $dbh->prepare ($date);
 $sth_date_1->execute;

while (my $x = $sth_date_1->fetchrow_hashref()) {
	$wk_st_date_key = $x->{WEEK_ST_DATE_KEY};
	$wk_en_date_key = $x->{WEEK_END_DATE_KEY};
	$wk_number = $x->{WEEK_NUMBER_THIS_YEAR};
	$as_of = $x->{DATE_FLD};
}

# $wk_st_date_key = 637; #TEST
# $wk_en_date_key = 643; #TEST

$date_2 = qq{ 
SELECT DATE_KEY1, TO_CHAR(DATE_FLD1, 'DD Mon YYYY') DATE_FLD1, TO_CHAR(DATE_FLD_LY1, 'DD Mon YYYY') DATE_FLD_LY1, DATE_KEY2, TO_CHAR(DATE_FLD2, 'DD Mon YYYY') DATE_FLD2, TO_CHAR(DATE_FLD_LY2, 'DD Mon YYYY') DATE_FLD_LY2, DATE_KEY3, DATE_FLD3, MONTH_ST_DATE_KEY, MONTH_END_DATE_KEY FROM
	(SELECT DATE_KEY AS DATE_KEY1, DATE_FLD AS DATE_FLD1, DATE_FLD_LY AS DATE_FLD_LY1
	FROM DIM_DATE WHERE DATE_KEY = $wk_st_date_key),
	(SELECT DATE_KEY AS DATE_KEY2, DATE_FLD AS DATE_FLD2, DATE_FLD_LY AS DATE_FLD_LY2
	FROM DIM_DATE WHERE DATE_KEY = $wk_en_date_key),
    (SELECT DATE_KEY AS DATE_KEY3, DATE_FLD AS DATE_FLD3, MONTH_ST_DATE_KEY, MONTH_END_DATE_KEY
	FROM DIM_DATE_PRL WHERE TO_CHAR(DATE_FLD, 'DD Mon YYYY') = '$as_of')
 };

my $sth_date_2 = $dbh->prepare ($date_2);
 $sth_date_2->execute;
 
while (my $x = $sth_date_2->fetchrow_hashref()) {
	$wk_st_date_fld = $x->{DATE_FLD1};
	$wk_en_date_fld = $x->{DATE_FLD2};
	$wk_st_date_fld_ly = $x->{DATE_FLD_LY1};
	$wk_en_date_fld_ly = $x->{DATE_FLD_LY2};
	$mo_st_date_key = $x->{MONTH_ST_DATE_KEY};
	$mo_en_date_key = $x->{DATE_KEY3};
}

# $mo_st_date_key = 274; #TEST
# $mo_en_date_key = 274; #TEST

$date_3 = qq{ 
SELECT DATE_KEY1, TO_CHAR(DATE_FLD1, 'DD Mon YYYY') DATE_FLD1, DATE_KEY_LY1, TO_CHAR(DATE_FLD_LY1, 'DD Mon YYYY') DATE_FLD_LY1, 
	   DATE_KEY2, TO_CHAR(DATE_FLD2, 'DD Mon YYYY') DATE_FLD2, DATE_KEY_LY2, TO_CHAR(DATE_FLD_LY2, 'DD Mon YYYY') DATE_FLD_LY2 FROM
	(SELECT DATE_KEY AS DATE_KEY1, DATE_FLD AS DATE_FLD1, DATE_KEY_LY AS DATE_KEY_LY1, DATE_FLD_LY AS DATE_FLD_LY1
	FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_st_date_key),
	(SELECT DATE_KEY AS DATE_KEY2, DATE_FLD AS DATE_FLD2, DATE_KEY_LY AS DATE_KEY_LY2, DATE_FLD_LY AS DATE_FLD_LY2
	FROM DIM_DATE_PRL WHERE DATE_KEY = $mo_en_date_key)
 };

my $sth_date_3 = $dbh->prepare ($date_3);
 $sth_date_3->execute;
 
while (my $x = $sth_date_3->fetchrow_hashref()) {
	$mo_st_date_key_ly = $x->{DATE_KEY_LY1};
	$mo_en_date_key_ly = $x->{DATE_KEY_LY2};
	$mo_st_date_fld = $x->{DATE_FLD1};
	$mo_en_date_fld = $x->{DATE_FLD2};
	$mo_st_date_fld_ly = $x->{DATE_FLD_LY1};
	$mo_en_date_fld_ly = $x->{DATE_FLD_LY2};
}

$sth_date_1->finish();
$sth_date_2->finish();
$sth_date_3->finish();
