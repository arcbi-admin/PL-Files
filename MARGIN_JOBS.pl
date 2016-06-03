use Win32::Job;

my $job1 = Win32::Job->new;
$job1->spawn( "cmd" , q{cmd /C "COMPILED_MARGIN_SALES_PERFORMANCE_v1.7_NORTHCEBU.pl pause"});
$job1->run(25000);

my $job2 = Win32::Job->new;
$job2->spawn( "cmd" , q{cmd /C "COMPILED_MARGIN_SALES_PERFORMANCE_v1.7_SOUTHCEBU.pl pause"});
$job2->run(25000);


my $job3 = Win32::Job->new;
$job3->spawn( "cmd" , q{cmd /C "COMPILED_MARGIN_SALES_PERFORMANCE_v1.7_NONCEBU.pl pause"});
$job3->run(25000);

my $job4 = Win32::Job->new;
$job4->spawn( "cmd" , q{cmd /C "COMPILED_MARGIN_SALES_PERFORMANCE_v1.7_NORTHLUZON.pl pause"});
$job4->run(25000);

my $job4 = Win32::Job->new;
$job4->spawn( "cmd" , q{cmd /C "COMPILED_MARGIN_SALES_PERFORMANCE_v1.7_SOUTHLUZON.pl pause"});
$job4->run(25000);

my $job4 = Win32::Job->new;
$job4->spawn( "cmd" , q{cmd /C "COMPILED_MARGIN_SALES_PERFORMANCE_v1.7_CENTRALLUZON.pl pause"});
$job4->run(25000);


#DRI KUTOB

#my $job5 = Win32::Job->new;
#$job5->spawn( "cmd" , q{cmd /C "BI_SALES_PERFORMANCE_v2.7xx.pl pause"});
#$job5->run(25000);










