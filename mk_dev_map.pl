#!/usr/bin/perl

use strict;

my %DeviceMap;
my $sar2xls_string = "IO - ";
my $charsleft;
my $prefix;

sub AddToDeviceMap {
	my ($DevString, $RegEx, $Prefix) = @_;
	my ($Device, %ThisDeviceMap);
	for $Device (`ls -l /dev/$DevString 2>/dev/null`){
		my @fields = split " ", $Device;
		my $Perms = $fields[0];
		next if substr ($Perms, 0, 1) eq "l";
		my $Major = $fields[4];
		chop $Major;
		my $Minor = $fields[5];
		my $DevNode = $fields[$#fields];
		$charsleft = 31 - length($sar2xls_string) - length($Prefix);
		$DevNode =~ /$RegEx/;
		if (length($1) > $charsleft) {
			$charsleft -= 3;
			$DevString = "..." . substr $1,-$charsleft,$charsleft;
		} else {
			$DevString = $1;
		}
		$ThisDeviceMap{"dev$Major-$Minor"} = "$Prefix$DevString";
	}
	return %ThisDeviceMap;
}

%DeviceMap = (%DeviceMap, AddToDeviceMap("sd\*", qr|\/dev\/([\w-]+)|, "dev~"));
%DeviceMap = (%DeviceMap, AddToDeviceMap("md\*", qr|\/dev\/([\w-]+)|, "dev~"));
%DeviceMap = (%DeviceMap, AddToDeviceMap("emcpower\*", qr|\/dev\/([\w-]+)|, "emcpower~"));
%DeviceMap = (%DeviceMap, AddToDeviceMap("mapper/\*", qr|\/dev\/mapper\/([\w-]+)|, "mapper~"));
%DeviceMap = (%DeviceMap, AddToDeviceMap("cciss/\*", qr|\/dev\/cciss\/([\w-]+)|, "cciss~"));

if (-e '/etc/init.d/oracleasm'){
        for my $asmline (`for i in \`service oracleasm listdisks\`; do service oracleasm querydisk -d \$i;done`){
                my ($junk1, $asmvolname, $junk2, $junk3, $junk4, $junk5, $junk6, $junk7, $junk8, $DevNode) = split " ", ;
                chop $asmvolname;
                $asmvolname = substr ($asmvolname, 1);
                $DevNode =~ /\[(\d{1,3}),(\d{1,3})/;
                my $Major = $1;
                my $Minor = $2;
                $DeviceMap{"dev$Major-$Minor"} = "ASM~$asmvolname";
        }
}

my $hostname = `hostname -s`;
chomp $hostname;
open (OUTFILE, ">${hostname}_dev_map");
print "Writing to ${hostname}_dev_map...\n";
print OUTFILE "[$hostname]\n";
foreach my $Device (sort keys %DeviceMap){
	print "$Device = $DeviceMap{$Device}\n";
	print OUTFILE "$Device = $DeviceMap{$Device}\n";
}
close (OUTFILE);

