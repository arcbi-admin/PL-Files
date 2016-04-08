use DBI;
use DBD::Oracle qw(:ora_types);

$hostname = "10.128.0.220";
$sid = "METROBIP";
$port = '1521';
$uid = 'ARCMA';
$pw = 'arcma';

$dbh = DBI->connect("dbi:Oracle:host=$hostname;sid=$sid;port=$port", $uid , $pw, { RaiseError => 1, AutoCommit => 0 }) or die "Unable to connect: $DBI::errstr";