
use Text::Iconv;
my $converter = Text::Iconv -> new ("utf-8", "windows-1251");

# Text::Iconv is not really required.
# This can be any object with the convert method. Or nothing.

use Spreadsheet::XLSX;

my $filename = shift @ARGV;
die "File $filename not found" unless -e $filename;

my $excel = Spreadsheet::XLSX -> new ($filename, $converter);
my $prev_line = "LINE";
foreach my $sheet (@{$excel -> {Worksheet}}) {
    
    printf("Sheet: %s\n", $sheet->{Name});
    
    $sheet -> {MaxRow} ||= $sheet -> {MinRow};
    
    foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}) {
        my $line = "LINE $row\n";
	if ($line ne $prev_line) {
		print $line;
		$prev_line = $line;
	}
        $sheet -> {MaxCol} ||= $sheet -> {MinCol};
        
        foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol}) {
            my $cell = $sheet -> {Cells} [$row] [$col];
            print "  COL $col:", filter($cell -> {Val}), "\n" if $cell;
        }
    }
}

sub filter {
    my $s = shift;
    $s =~ s/\n/\\n/g;
    $s =~ s/&amp;/\&/g;
    $s =~ s/&gt;/>/g;
    $s =~ s/&lt;/</g;
    return $s;
}
