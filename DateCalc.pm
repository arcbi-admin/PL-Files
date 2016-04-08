use Date::Calc qw( Today Add_Delta_Days );

($year,$month,$day) = Today(); #today's date
($year,$month,$day) = Add_Delta_Days($year,$month,$day,-1);

if ($day <= 9){
	if ($month <= 9){
		$b_date = ("$year-0$month-0$day");
	}
	else{
		$b_date = ("$year-$month-0$day");
	}
}
else{
	if ($month <= 9 ){
		$b_date = ("$year-0$month-$day");
	}
	else{
		$b_date = ("$year-$month-$day");
	}	
}

($year,$month,$day) = Add_Delta_Days($year,$month,$day,-6);

if ($day <= 9){
	if ($month <= 9){
		$st_date = ("$year-0$month-0$day");
	}
	else{
		$st_date = ("$year-$month-0$day");
	}
}
else{
	if ($month <= 9 ){
		$st_date = ("$year-0$month-$day");
	}
	else{
		$st_date = ("$year-$month-$day");
	}	
}

