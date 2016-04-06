use Date::Calc qw( Today Day_of_Week Add_Delta_Days
                     Day_of_Week_to_Text Date_to_Text Decode_Date_US );

  $searching_dow = 7; # 7 = Sunday

  @today = Today();

  $current_dow = Day_of_Week(@today);

  if ($searching_dow == $current_dow)
  {
      @prev = Add_Delta_Days(@today,-7);
      @next = Add_Delta_Days(@today,+7);
  }
  else
  {
      if ($searching_dow > $current_dow)
      {
          @next = Add_Delta_Days(@today,$searching_dow - $current_dow);
          @prev = Add_Delta_Days(@next,-7);
      }
      else
      {
          @prev = Add_Delta_Days(@today,$searching_dow - $current_dow);
          @next = Add_Delta_Days(@prev,+7);
      }
  }

  $dow = Day_of_Week_to_Text($searching_dow);

  $wtd_end = Day_of_Week(@prev);
  
  print "Today is:      ", ' ' x length($dow), Date_to_Text(@today), "\n";
  print "Last $dow was:     ", $wtd_end,  "\n";
  
  
  