use DBI;
use DBD::Oracle qw(:ora_types);
use Date::Calc qw( Today Day_of_Week Month_to_Text Week_Number );
#use DBConnector;
# use DBConnectorCommodus;

$hostname = "10.128.0.220";
$sid = "METROBIP";

$port = '1521';
$uid = 'ARCMA';
$pw = 'arcma';

$dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw, { RaiseError => 1, AutoCommit => 0 }) or die "Unable to connect: $DBI::errstr";

$y_st_date_key = 2430 ; #JAN 1 2013
$y_st_date_key_ly = 2064 ; #JAN 1 2012

$y_st_date_key2012 = 2064 ; #2012

($year,$month,$day) = Today(); #date today
  
$week = int(($day + Day_of_Week($year,$month,1) - 2) / 7) + 1; #week number this week

$week = Week_Number($year,$month,$day); #computes week number of the year

$reg_week = $week - 1 ; #variable to hold week number for the last full week

if ($month eq 1) { $mo = 'Jan'; }
elsif ($month eq 2) { $mo = 'Feb';  }
elsif ($month eq 3) { $mo = 'Mar';  }
elsif ($month eq 4) { $mo = 'Apr';  }
elsif ($month eq 5) { $mo = 'May';  }
elsif ($month eq 6) { $mo = 'Jun';  }
elsif ($month eq 7) { $mo = 'Jul';  }
elsif ($month eq 8) { $mo = 'Aug';  }
elsif ($month eq 9) { $mo = 'Sep';  }
elsif ($month eq 10) { $mo = 'Oct';  }
elsif ($month eq 11) { $mo = 'Nov';  }
elsif ($month eq 12) { $mo = 'Dec';  }

$year_ly =  $year - 1;

$sql_date = ("$day-$mo-$year");

$query = "SELECT * FROM (SELECT DATE_KEY, DATE_FLD, MONTH_NAME, YEAR_NAME FROM DIM_DATE WHERE WEEK_NUMBER_THIS_YEAR = (SELECT WEEK_NUMBER_THIS_YEAR FROM DIM_DATE WHERE DATE_FLD = '$sql_date')-1 AND DAY_IN_WEEK = 7  AND YEAR_NAME = (SELECT YEAR_NAME FROM DIM_DATE WHERE DATE_FLD= '$sql_date'))A, (SELECT DATE_KEY AS DATE_KEY_LY, DATE_FLD AS DATE_FLD_LY FROM DIM_DATE WHERE DATE_FLD = (SELECT (SUBSTR(K,1,6) || - $year_ly) FROM(SELECT DATE_FLD AS K FROM DIM_DATE WHERE WEEK_NUMBER_THIS_YEAR = (SELECT WEEK_NUMBER_THIS_YEAR FROM DIM_DATE WHERE DATE_FLD = '$sql_date')-1 AND DAY_IN_WEEK = 7  AND YEAR_NAME = (SELECT YEAR_NAME FROM DIM_DATE WHERE DATE_FLD= '$sql_date'))))B";

# $query = "SELECT * FROM(SELECT DATE_KEY, DATE_FLD, MONTH_NAME, YEAR_NAME FROM DIM_DATE WHERE DATE_FLD = '31-JUL-2013')A, (SELECT DATE_KEY AS DATE_KEY_LY, DATE_FLD AS DATE_FLD_LY FROM DIM_DATE WHERE DATE_FLD = '31-JUL-2012')B";   ##ENDING

$handler = $dbh->prepare($query); 
$handler->execute();

while ($y = $handler->fetchrow_hashref()) {
		$as_of_date = $y->{DATE_FLD};
		$end_date_key = $y->{DATE_KEY};
		$end_date_key_ly = $y->{DATE_KEY_LY};
		$month = $y->{MONTH_NAME};
}


# $query2 = "SELECT * FROM (SELECT DATE_KEY AS DATE_KEY2,DATE_FLD AS DATE_FLD2 FROM DIM_DATE WHERE DATE_KEY = (SELECT DATE_KEY FROM DIM_DATE WHERE DATE_FLD = (SELECT VALUE FROM ADMIN_ETL_SUMMARY)) - ((SELECT SUBSTR(K,1,2) FROM(SELECT VALUE AS K FROM ADMIN_ETL_SUMMARY))-1))A, (SELECT DATE_KEY AS DATE_KEY2_LY, DATE_FLD AS DATE_FLD2_LY FROM DIM_DATE WHERE DATE_FLD = (SELECT (SUBSTR(K,1,6) || - $year_ly) FROM(SELECT DATE_FLD AS K FROM DIM_DATE WHERE DATE_KEY = (SELECT DATE_KEY FROM DIM_DATE WHERE DATE_FLD = (SELECT VALUE FROM ADMIN_ETL_SUMMARY)) - ((SELECT SUBSTR(K,1,2) FROM(SELECT VALUE AS K FROM ADMIN_ETL_SUMMARY))-1))))B";

# $query2 = "SELECT * FROM (SELECT DATE_KEY AS DATE_KEY2,DATE_FLD AS DATE_FLD2 FROM DIM_DATE WHERE DATE_FLD = '01-JUL-2013')A, (SELECT DATE_KEY AS DATE_KEY2_LY,DATE_FLD AS DATE_FLD2_LY FROM DIM_DATE WHERE DATE_FLD = '01-JUL-2012')B";   ##MONTHLY BEGINNING

# $handler2 = $dbh->prepare($query2); 
# $handler2->execute();

# while ($y2 = $handler2->fetchrow_hashref()) {
		# $w_st_date_key = $y2->{DATE_KEY2}; 
		# $m_st_date_key = $y2->{DATE_KEY2};
		# $w_st_date_key_ly = $y2->{DATE_KEY2_LY}; 
		# $m_st_date_key_ly = $y2->{DATE_KEY2_LY};
# }

# print "SQL DATE: ".$sql_date."\n"; 
# print "AS OF DATE: ".$as_of_date."\n";
# print "WK STRT KEY: ".$w_st_date_key."\n";
# print "MO STRT KEY: ".$m_st_date_key."\n";
# print "END DTE KEY: ".$end_date_key."\n";
# print "AS OF WEEK: ".$as_of_week."\n";
# print "MONTH: ".$month."\n";

$date_query = "SELECT DATE_KEY, DATE_FLD, MONTH_ST_DATE_KEY, MONTH_END_DATE_KEY FROM DIM_DATE WHERE DATE_KEY = (SELECT MAX(DATE_KEY) FROM DIM_DATE WHERE DATE_KEY >= $y_st_date_key2012 AND DATE_KEY <= (SELECT DATE_KEY  FROM DIM_DATE WHERE DATE_FLD = (SELECT MAX(ETL_TO_DATE)  FROM INC_ETL_HISTORY WHERE BATCH_ID = 8 AND STATUS = 1)) AND DAY_IN_MONTH IN ( 28 , 35 )) ";
$date_handler = $dbh->prepare($date_query); 
$date_handler->execute();

while ($x = $date_handler->fetchrow_hashref()) {
	$month_date = $x->{DATE_FLD};
	$month_start = $x->{MONTH_ST_DATE_KEY};
	$month_end = $x->{MONTH_END_DATE_KEY};
}

# print "MONTH DATE: ".$month_date."\n";
# print "MO ST KEY: ".$month_start."\n";
# print "MO END KEY: ".$month_end."\n";

$handler->finish();   
$handler2->finish();   
$date_handler->finish();

sub mydate {
@date = split /-/, $as_of_date;
if (@date[1] eq 'JAN') { $cardinal = '01'; $sintenel = 'Jan'; }
elsif (@date[1] eq 'FEB') { $cardinal = '02'; $sintenel = 'Feb';  }
elsif (@date[1] eq 'MAR') { $cardinal = '03'; $sintenel = 'Mar';  }
elsif (@date[1] eq 'APR') { $cardinal = '04'; $sintenel = 'Apr';  }
elsif (@date[1] eq 'MAY') { $cardinal = '05'; $sintenel = 'May';  }
elsif (@date[1] eq 'JUN') { $cardinal = '06'; $sintenel = 'Jun';  }
elsif (@date[1] eq 'JUL') { $cardinal = '07'; $sintenel = 'Jul';  }
elsif (@date[1] eq 'AUG') { $cardinal = '08'; $sintenel = 'Aug';  }
elsif (@date[1] eq 'SEP') { $cardinal = '09'; $sintenel = 'Sep';  }
elsif (@date[1] eq 'OCT') { $cardinal = '10'; $sintenel = 'Oct';  }
elsif (@date[1] eq 'NOV') { $cardinal = '11'; $sintenel = 'Nov';  }
elsif (@date[1] eq 'DEC') { $cardinal = '12'; $sintenel = 'Dec';  }
printf "As of: ".$cardinal."/".@date[0]."/20".@date[2]."\n"; 
}





