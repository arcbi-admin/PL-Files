use Win32::Job;

#my $job1 = Win32::Job->new;
#$job1->spawn( "cmd" , q{cmd /C "BI_SALES_PERFORMANCE_v1.8.pl pause"});
#$job1->run(25000);

my $job2 = Win32::Job->new;
$job2->spawn( "cmd" , q{cmd /C "BI_SALES_PERFORMANCE_v2.8.pl pause"});
$job2->run(25000);

#my $job3 = Win32::Job->new;
#$job3->spawn( "cmd" , q{cmd /C "BI_SALES_PERF_CON_v1.3.pl pause"});
#$job3->run(25000);

#my $job4 = Win32::Job->new;
#$job4->spawn( "cmd" , q{cmd /C "BI_SALES_PERF_OTR_v1.3.pl pause"});
#$job4->run(25000);

#my $job5 = Win32::Job->new;
#$job5->spawn( "cmd" , q{cmd /C "BI_SALES_PERFORMANCE_v2.7xx.pl pause"});
#$job5->run(25000);










