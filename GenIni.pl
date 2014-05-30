#!/usr/bin/perl -w
#-------------------------------------------------------------------------------
# Author:   Peter McGowan, Red Hat (pemcg@redhat.com)
#
# Description:  genini.pl
#
#
# Revision History
#
# Original  0.1     28-Feb-2012     PEMcG   Original version
# Version:  0.2     06-Mar-2012     PEMcG   Use '::' rather than ":' as separators in Data section of the ini file
#           0.3     07-Mar-2012     PEMcG   Significant re-write to prompt for nearly everything required 

use constant VERSION => 0.3;
#
#
#-------------------------------------------------------------------------------

use strict;
use English;
use Getopt::Std;
use File::Spec;
use Spreadsheet::ParseExcel;
use Term::Prompt;
$Term::Prompt::MULTILINE_INDENT = '';
$Text::Wrap::columns = 132;

use constant MAX_SHEET_NAME_LENGTH => 23;
use constant FALSE => 0x0;
use constant TRUE => 0x1;

my $Slash = '/';        # Assume unix
$Slash = '\\' if $OSNAME eq "MSWin32";

#
# Sheet title, Graph title, Y-axis title and CellDivisionFactor for a given sar header in a sar2xls spreadsheet
#
my %GraphAttributes = (
    "% Allocated Disc Quota Entries" => ["Files %dquot-sz", "Percent Allocated Disc Quota Entries", "Percent", 1],
    "% Allocated Super Block Handlers" => ["Files %super-sz", "Percent Allocated Super Block Handlers", "Percent", 1],
    "% Bandwidth Utilisation" => ["IO DEVICE %util", "DEVICE Bandwidth Utilisation", "Percent", 1],
    "% Cached Swap/Used Swap" => ["Swap %swpcad", "Percent Cached Swap/Used Swap", "Percent", 1],
    "% Efficiency of Page Reclaim" => ["Memory %vmeff", "Percentage Efficiency of Page Reclaim", "Percent", 1],
    "% Idle Waiting on I/O (IOwait)" => ["CPU %iowait", "CPU Percent Idle Waiting on I/O", "Percent", 1],
    "% Idle" => ["CPU %idle", "CPU Percent Idle", "Percent", 1],
    "% Memory Required for Current Workload" => ["Memory %commit", "Percent Memory Required for Current Workload", "Percent", 1],
    "% Memory Used" => ["Memory %memused", "Percent Memory Used", "Percent", 1],
    "% Queued RT Signals" => ["Files %rtsig-sz", "Percent Queued RT Signals", "Percent", 1],
    "% Swap Used" => ["Swap %swpused", "Percent Swap Used", "Percent", 1],
    "% System Level (excl. interrupts)" => ["CPU %sys", "CPU Percent System Level (excl. interrupts)", "Percent", 1],
    "% System Level (inc. interrupts)" => ["CPU %system", "CPU Percent System Level (inc. interrupts)", "Percent", 1],
    "% System Level" => ["CPU %system", "CPU Percent System Level", "Percent", 1],
    "% Time Involuntary Waiting on Another vCPU" => ["CPU %steal", "CPU Percent Time Involuntary Waiting on Another vCPU", "Percent", 1],
    "% Time Servicing HW Interrupts" => ["CPU %irq", "CPU Percent Time Servicing HW Interrupts", "Percent", 1],
    "% Time Servicing SW Interrupts" => ["CPU %soft", "CPU Percent Time Servicing SW Interrupts", "Percent", 1],
    "% Time running a vCPU" => ["CPU %guest", "CPU Percent Time running a vCPU", "Percent", 1],
    "% Used File Handles" => ["Files file-nr", "Percent Used File Handles", "Percent", 1],
    "% User Level (nice)" => ["CPU %nice", "CPU Percent User Level (nice)", "Percent", 1],
    "% User Level excl. vCPU" => ["CPU %usr", "CPU Percent User Level excl. vCPU", "Percent", 1],
    "% User Level inc. vCPU" => ["CPU %user", "CPU Percent User Level inc. vCPU", "Percent", 1],
    "% User Level" => ["CPU %user", "CPU Percent User Level", "Percent", 1],
    "% Utilisation" => ["CPU %Utilisation", "CPU Percent Utilisation", "Percent", 1],
    "Active Memory Pages" => ["Memory activepg", "Active Memory Pages", "Pages", 1],
    "Add'l Buffer Pages/Sec" => ["Memory bufpgps", "Additional Buffer Pages/Sec", "Pages/Sec", 1],
    "Add'l Cache Pages/Sec" => ["Memory campgps", "Additional Cache Pages/Sec", "Pages/Sec", 1],
    "Add'l Pages Shared/Sec" => ["Memory shmpgps", "Additional Pages Shared/Sec", "Pages/Sec", 1],
    "Allocated Disc Quota Entries" => ["Files %dquot-sz", "Allocated Disc Quota Entries", "Entries", 1],
    "Average I/O Service Time (ms)" => ["IO DEVICE svctm", "DEVICE Average I/O Service Time", "mSec", 1],
    "Average I/O Svc Time (unreliable)" => ["IO DEVICE svctm", "DEVICE Average I/O Service Time", "mSec", 1],
    "Average I/O Time inc Wait(ms)" => ["IO DEVICE await", "DEVICE Average I/O Time including Wait", "mSec", 1],
    "Average Queue Length" => ["IO DEVICE avgqu-sz", "DEVICE Average Queue Length", "Queued I/Os", 1],
    "Average Request Size (Sectors)" => ["IO DEVICE avgrq-sz", "DEVICE Average Request Size", "KBytes", 2],
    "Bad Pkts Recv'd/Sec" => ["Net Err DEVICE rxerrps", "DEVICE Bad Packets Received/Sec", "Packets", 1],
    "Bad RPC Requests Recv'd/Sec" => ["NFS Server badcallps", "Bad RPC Requests Received/Sec", "Requests", 1],
    "Bad TCP Segments Recv'd/Sec" => ["TCPv4 Errors isegerrps", "Bad TCP Segments Received/Sec", "Segments", 1],
    "Blocks Read/Sec" => ["Total IO breadps", "Blocks Read/Sec", "Blocks", 1],
    "Blocks Tranferred/Sec" => ["Total IO blksps", "Blocks Tranferred/Sec", "Blocks", 1],
    "Blocks Written/Sec" => ["Total IO bwrtnps", "Blocks Written/Sec", "Blocks", 1],
    "Bytes Recv'd/Sec" => ["Network DEVICE rxbytps", "DEVICE MBytes Received/Sec", "MBytes", 1048576],
    "Bytes Trans'd/Sec" => ["Network DEVICE txbytps", "DEVICE MBytes Transmitted/Sec", "MBytes", 1048576],
    "CPU Clock Frequency (MHz)" => ["Power Mgmt MHz", "CPU Clock Frequency", "MHz", 1],
    "Cached Swap KB" => ["Swap kbswpcad", "Cached Swap", "KBytes", 1],
    "Child Process % System Level" => ["Child Process %csystem", "Child Process % System Level", "Percent", 1],
    "Child Process % User Level" => ["Child Process %cuser", "Child Process % User Level", "Percent", 1],
    "Child Process Major Faults/Sec" => ["Child Process cmajfltps", "Child Process Major Faults/Sec", "Major Faults", 1],
    "Child Process Minor Faults/Sec" => ["Child Process cminfltps", "Child Process Minor Faults/Sec", "Minor Faults", 1],
    "Child Process Pages Swapped Out/Sec" => ["Child Process cnswapps", "Child Process Pages Swapped Out/Sec", "Pages", 1],
    "Collisions/Sec" => ["Net Err DEVICE collps", "DEVICE Collisions/Sec", "Collisions", 1],
    "Compressed Pkts Recv'd/Sec" => ["Network DEVICE rxcmpps", "DEVICE Compressed Packets Received/Sec", "Packets", 1],
    "Compressed Pkts Trans'd/Sec" => ["Network DEVICE txcmpps", "DEVICE Compressed Packets Transmitted/Sec", "Packets", 1],
    "Context Switches/Sec" => ["Context Switches", "Context Switches/Sec", "Context Switches", 1],
    "Device Transfers/Sec" => ["IO DEVICE tps", "DEVICE I/O Transfers/Sec", "Transfers", 1],
    "D'gram Fragments Created/Sec" => ["IPv4 fragcrtps", "Datagram Fragments Created/Sec", "Fragments", 1],
    "D'grams Discarded No Route/Sec" => ["IPv4 Errors onortps", "Datagrams Discarded No Route/Sec", "Datagrams", 1],
    "D'grams Not Fragmented/Sec" => ["IPv4 Errors fragfps", "Datagrams Not Fragmented/Sec", "Datagrams", 1],
    "D'grams Successfully Fragmented/Sec" => ["IPv4 fragokps", "Datagrams Successfully Fragmented/Sec", "Datagrams", 1],
    "D'grams Successfully Reassembed/Sec" => ["IPv4 asmokps", "Datagrams Successfully Reassembed/Sec", "Datagrams", 1],
    "Data Cache KB" => ["Memory kbcached", "Data Cache MB", "MBytes", 1024],
    "Forwarded D'grams/Sec" => ["IPv4 fwddgmps", "Forwarded D'grams/Sec", "Datagrams", 1],
    "Fragments Needing Reassembly/Sec" => ["IPv4 asmrqps", "Fragments Needing Reassembly/Sec", "Fragments", 1],
    "Free Memory KB" => ["Memory kbmemfree", "Free Memory MB", "MBytes", 1024],
    "Free Swap Space KB" => ["Swap kbswpfree", "Free Swap Space MB", "MBytes", 1024],
    "ICMP Address Mask Replies Recv'd/Sec" => ["ICMPv4 iadrmkrps", "ICMP Address Mask Replies Received/Sec", "Replies", 1],
    "ICMP Address Mask Replies Trans'd/Sec" => ["ICMPv4 oadrmkrps", "ICMP Address Mask Replies Trans'd/Sec", "Replies", 1],
    "ICMP Address Mask Requests Recv'd/Sec" => ["ICMPv4 iadrmkps", "ICMP Address Mask Requests Received/Sec", "Requests", 1],
    "ICMP Address Mask Requests Trans'd/Sec" => ["ICMPv4 iadrmkrps", "ICMP Address Mask Replies Received/Sec", "Requests", 1],
    "ICMP Dest. Unreach Msgs Recv'd/Sec" => ["ICMPv4 Errors idstunrps", "ICMP Dest. Unreach Msgs Received/Sec", "Messages", 1],
    "ICMP Dest. Unreach Msgs Trans'd/Sec" => ["ICMPv4 Errors odstunrps", "ICMP Dest. Unreach Msgs Trans'd/Sec", "Messages", 1],
    "ICMP Echo Replies Recv'd/Sec" => ["ICMPv4 iechrps", "ICMP Echo Replies Received/Sec", "Replies", 1],
    "ICMP Echo Replies Trans'd/Sec" => ["ICMPv4 oechrps", "ICMP Echo Replies Transmitted/Sec", "Replies", 1],
    "ICMP Echo Requests Recv'd/Sec" => ["ICMPv4 iechps", "ICMP Echo Requests Received/Sec", "Requests", 1],
    "ICMP Echo Requests Trans'd/Sec" => ["ICMPv4 oechps", "ICMP Echo Requests Transmitted/Sec", "Requests", 1],
    "ICMP Msgs Not Trans'd/Sec" => ["ICMPv4 Errors oerrps", "ICMP Msgs Not Transmitted/Sec", "Messages", 1],
    "ICMP Msgs Recv'd/Sec" => ["ICMPv4 imsgps", "ICMP Msgs Recv'd/Sec", "Messages", 1],
    "ICMP Msgs Trans'd (attempted)/Sec" => ["ICMPv4 omsgps", "ICMP Msgs Trans'd (attempted)/Sec", "Messages", 1],
    "ICMP Msgs With Errors Recv'd/Sec" => ["ICMPv4 Errors ierrps", "ICMP Msgs With Errors Received/Sec", "Messages", 1],
    "ICMP Parameter Problem Msgs Recv'd/Sec" => ["ICMPv4 Errors iparmpbps", "ICMP Parameter Problem Msgs Recv'd/Sec", "Messages", 1],
    "ICMP Parameter Problem Msgs Trans'd/Sec" => ["ICMPv4 Errors oparmpbps", "ICMP Parameter Problem Msgs Transmitted/Sec", "Messages", 1],
    "ICMP Redirect Msgs Recv'd/Sec" => ["ICMPv4 Errors iredirps", "ICMP Redirect Msgs Received/Sec", "Messages", 1],
    "ICMP Redirect Msgs Trans'd/Sec" => ["ICMPv4 Errors oredirps", "ICMP Redirect Msgs Transmitted/Sec", "Messages", 1],
    "ICMP Source Quench Msgs Recv'd/Sec" => ["ICMPv4 Errors isrcqps", "ICMP Source Quench Msgs Received/Sec", "Messages", 1],
    "ICMP Source Quench Msgs Trans'd/Sec" => ["ICMPv4 Errors osrcqps", "ICMP Source Quench Msgs Transmitted/Sec", "Messages", 1],
    "ICMP Time Exceeded Msgs Recv'd/Sec" => ["ICMPv4 Errors itmexps", "ICMP Time Exceeded Msgs Received/Sec", "Messages", 1],
    "ICMP Time Exceeded Msgs Trans'd/Sec" => ["ICMPv4 Errors otmexps", "ICMP Time Exceeded Msgs Transmitted/Sec", "Messages", 1],
    "ICMP Timestamp Replies Recv'd/Sec" => ["ICMPv4 itmrps", "ICMP Timestamp Replies Received/Sec", "Replies", 1],
    "ICMP Timestamp Replies Trans'd/Sec" => ["ICMPv4 otmrps", "ICMP Timestamp Replies Transmitted/Sec", "Replies", 1],
    "ICMP Timestamp Requests Recv'd/Sec" => ["ICMPv4 itmps", "ICMP Timestamp Requests Received/Sec", "Requests", 1],
    "ICMP Timestamp Requests Trans'd/Sec" => ["ICMPv4 otmps", "ICMP Timestamp Requests Transmitted/Sec", "Requests", 1],
    "IP Fragments in Use" => ["Socket ip-frag", "IP Fragments in Use", "Fragments", 1],
    "IP Reassembly Failures/Sec" => ["IPv4 Errors asmfps", "IP Reassembly Failures/Sec", "Reassemblies", 1],
    "Inactive Clean Pages" => ["Memory inaclnpg", "Inactive Clean Pages", "Pages", 1],
    "Inactive Dirty Pages" => ["Memory inadtypg", "Inactive Dirty Pages", "Pages", 1],
    "Inactive Target Pages" => ["Memory inatarpg", "Inactive Target Pages", "Pages", 1],
    "Input D'gram IP Header Errors/Sec" => ["IPv4 Errors ihdrerrps", "Input Datagram IP Header Errors/Sec", "Datagrams", 1],
    "Input D'gram IP Header Invalid Address/Sec" => ["IPv4 Errors iadrerrps", "Input Datagram IP Header Invalid Address/Sec", "Datagrams", 1],
    "Input D'gram Unknown Protocol/Sec" => ["IPv4 Errors iukwnprps", "Input Datagram Unknown Protocol/Sec", "Datagrams", 1],
    "Input D'grams Discarded/Sec" => ["IPv4 Errors idiscps", "Input Datagrams Discarded/Sec", "Datagrams", 1],
    "Input D'grams Successfully Delivered/Sec" => ["IPv4 idelps", "Input Datagrams Successfully Delivered/Sec", "Datagrams", 1],
    "Input D'grams/Sec" => ["IPv4 irecps", "Input Datagrams/Sec", "Datagrams", 1],
    "Interrupts/Sec" => ["Interrupts", "Interrupts/Sec", "Interrupts", 1],
    "KB Paged In/Sec" => ["Paging pgpginps", "KBytes Paged In/Sec", "KBytes", 1],
    "KB Paged Out/Sec" => ["Paging pgpgoutps", "KBytes Paged Out/Sec", "KBytes", 1],
    "KB Required for Current Workload" => ["Memory kbcommit", "MB Required for Current Workload", "MBytes", 1024],
    "KBytes Recv'd/Sec" => ["Network DEVICE rxkBps", "DEVICE MBytes Received/Sec", "MBytes", 1024],
    "KBytes Trans'd/Sec" => ["Network DEVICE txkBps", "DEVICE MBytes Transmitted/Sec", "MBytes", 1024],
    "Kernel Buffers KB" => ["Memory kbbuffers", "Kernel Buffers (MBytes)", "MBytes", 1024],
    "Locally Originated D'grams/Sec" => ["IPv4 orqps", "Locally Originated Datagrams/Sec", "Datagrams", 1],
    "Major Faults/Sec" => ["Paging majfltps", "Major Faults/Sec", "Faults", 1],
    "Memory Pages Freed/Sec" => ["Memory frmpgps", "Memory Pages Freed/Sec", "Pages", 1],
    "Minor Faults/Sec" => ["Process minfltps", "Minor Faults/Sec", "Faults", 1],
    "Multicast Pkts Recv'd/Sec" => ["Network DEVICE rxmcstps", "DEVICE Multicast Packets Received/Sec", "Packets", 1],
    "Network Packets Recv'd/Sec" => ["NFS Server packetps", "Network Packets Received/Sec", "Packets", 1],
    "No. of Processes Blocked on I/O" => ["Load Average blocked", "Number of Processes Blocked on I/O", "Processes", 1],
    "No. of Processes in List" => ["Load Average plist-sz", "Number of Processes in List", "Processes", 1],
    "Number of Pseudo-Terminals in Use" => ["Files pty-nr", "Number of Pseudo-Terminals in Use", "Pseudo-Terminals", 1],
    "Output D'grams Discarded/Sec" => ["IPv4 Errors odiscps", "Output Datagrams Discarded/Sec", "Datagrams", 1],
    "Pages Addeed to Free List/Sec" => ["Paging pgfreeps", "Pages Addeed to Free List/Sec", "Pages", 1],
    "Pages Reclaimed from Cache/Sec" => ["Paging pgstealps", "Pages Reclaimed from Cache/Sec", "Pages", 1],
    "Pages Scanned Directly/Sec" => ["Paging pgscandps", "Pages Scanned Directly/Sec", "Pages", 1],
    "Pages Scanned by kswapd/Sec" => ["Paging pgscankps", "Pages Scanned by kswapd/Sec", "Pages", 1],
    "Pages Swapped In/Sec" => ["Swapping pswpinps", "Pages Swapped In/Sec", "Pages", 1],
    "Pages Swapped Out/Sec" => ["Swapping pswpoutps", "Pages Swapped Out/Sec", "Pages", 1],
    "Pkts Recv'd/Sec" => ["Network DEVICE rxpckps", "DEVICE Packets Received/Sec", "Packets", 1],
    "Pkts Trans'd/Sec" => ["Network DEVICE txpckps", "DEVICE Packets Transmitted/Sec", "Packets", 1],
    "Process Pages Swapped Out/Sec" => ["Process nswapps", "Process Pages Swapped Out/Sec", "Pages", 1],
    "Processes Created/Sec" => ["Process procps", "Processes Created/Sec", "Processes", 1],
    "Queued RT Signals" => ["Files rtsig-sz", "Queued RT Signals", "Signals", 1],
    "RPC Access Calls Recv'd/Sec" => ["NFS Server saccessps", "RPC Access Calls Received/Sec", "Calls", 1],
    "RPC Access Requests Made/Sec" => ["NFS Client accessps", "RPC Access Requests Made/Sec", "Requests", 1],
    "RPC Getattr Calls Recv'd/Sec" => ["NFS Server sgetattps", "RPC Getattr Calls Received/Sec", "Calls", 1],
    "RPC Getattr Requests Made/Sec" => ["NFS Client getattps", "RPC Getattr Requests Made/Sec", "Requests", 1],
    "RPC Read Calls Recv'd/Sec" => ["NFS Server sreadps", "RPC Read Calls Received/Sec", "Calls", 1],
    "RPC Read Requests Made/Sec" => ["NFS Client readps", "RPC Read Requests Made/Sec", "Requests", 1],
    "RPC Requests Made/Sec" => ["NFS Client callps", "RPC Requests Made/Sec", "Requests", 1],
    "RPC Requests Recv'd/Sec" => ["NFS Server scallps", "RPC Requests Received/Sec", "Requests", 1],
    "RPC Retransmitted Requests Made/Sec" => ["NFS Client retransps", "RPC Retransmitted Requests Made/Sec", "Requests", 1],
    "RPC Write Calls Recv'd/Sec" => ["NFS Server swriteps", "RPC Write Calls Received/Sec", "Calls", 1],
    "RPC Write Requests Made/Sec" => ["NFS Client writeps", "RPC Write Requests Made/Sec", "Requests", 1],
    "Raw Sockets in Use" => ["Sockets rawsck", "Raw Sockets in Use", "Sockets", 1],
    "Read Transfers/Sec" => ["Total IO rtps", "Total IO Read Transfers/Sec", "Transfers", 1],
    "Recv FIFO Overrun Errors/Sec" => ["Net Err DEVICE rxfifops", "DEVICE Receive FIFO Overrun Errors/Sec", "Errors", 1],
    "Recv Frame Alignment Errors/Sec" => ["Net Err DEVICE rxframps", "DEVICE Receive Frame Alignment Errors/Sec", "Errors", 1],
    "Recv Pkts Dropped/Sec" => ["Net Err DEVICE rxdropps", "DEVICE Receive Packets Dropped/Sec", "Packets", 1],
    "Reply Cache Hits/Sec" => ["NFS Server hitps", "Reply Cache Hits/Sec", "Hits", 1],
    "Reply Cache Misses/Sec" => ["NFS Server missps", "Reply Cache Misses/Sec", "Misses", 1],
    "Run Queue Length" => ["Load Average runq-sz", "Run Queue Length", "Processes", 1],
    "Sectors Read/Sec" => ["IO DEVICE rd_secps", "DEVICE KBytes Read/Sec", "KBytes", 2],
    "Sectors Written/Sec" => ["IO DEVICE wr_secps", "DEVICE KBytes Written/Sec", "KBytes", 2],
    "Serial Line Breaks/Sec" => ["Serial Line brkps", "Serial Line Breaks/Sec", "Line Breaks", 1],
    "Serial Line Frame Errors/Sec" => ["Serial Line framerrps", "Serial Line Frame Errors/Sec", "Frame Errors", 1],
    "Serial Line Overruns/Sec" => ["Serial Line ovrunps", "Serial Line Overruns/Sec", "Overruns", 1],  
    "Serial Line Parity Errors/Sec" => ["Serial Line prtyerrps", "Serial Line Parity Errors/Sec", "Parity Errors", 1],
    "Serial Line Recv Ints/Sec" => ["Serial Line recvinps", "Serial Line Receive Interrupts/Sec", "Interrupts", 1],
    "Serial Line Trans Ints/Sec" => ["Serial Line xmtinps", "Serial Line Transmit Interrupts/Sec", "Interrupts", 1],
    "Shared Memory KB" => ["Memory kbmemshrd", "Shared Memory (MBytes)", "MBytes", 1024],
    "Super Block Handlers" => ["Files super-sz", "Super Block Handlers", "Handlers", 1],
    "System Load Avg. Last 15 Mins." => ["Load Average ldavg-15", "System Load Average Last 15 Minutes", "Load Average", 1],
    "System Load Avg. Last 5 Mins." => ["Load Average ldavg-5", "System Load Average Last 5 Minutes", "Load Average", 1],
    "System Load Avg. Last Min." => ["Load Average ldavg-1", "System Load Average Last Minute", "Load Average", 1],
    "TCP Active Opens/Sec" => ["TCPv4 activeps", "TCP Active Opens/Sec", "Opens", 1],
    "TCP Attempt Fails/Sec" => ["TCPv4 Errors atmptfps", "TCP Attempt Fails/Sec", "Failures", 1],
    "TCP Establish Resets/Sec" => ["TCPv4 Errors estresps", "TCP Establish Resets/Sec", "Resets", 1],
    "TCP Packets Recv'd/Sec" => ["NFS Server tcpps", "TCP Packets Received/Sec", "TCP Packets", 1],
    "TCP Passive Opens/Sec" => ["TCPv4 passiveps", "TCP Passive Opens/Sec", "TCP Opens", 1],
    "TCP RSTs Trans'd/Sec" => ["TCPv4 Errors orstsps", "TCP RSTs Transmitted/Sec", "TCP RSTs", 1],
    "TCP Segments Recv'd/Sec" => ["TCPv4 isegps", "TCP Segments Received/Sec", "TCP Segments", 1],
    "TCP Segments Retrans'd/Sec" => ["TCPv4 Errors retransps", "TCP Segments Retransmitted/Sec", "TCP Segments", 1],
    "TCP Segments Trans'd/Sec" => ["TCPv4 osegps", "TCP Segments Transmitted/Sec", "TCP Segments", 1],
    "TCP Sockets in TIME_WAIT State" => ["Socket tcp-tw", "TCP Sockets in TIME_WAIT State", "TCP Sockets", 1],
    "TCP Sockets in Use" => ["Socket tcpsck", "TCP Sockets in Use", "TCP Sockets", 1],
    "Total Page Faults/Sec" => ["Paging faultps", "Total Page Faults/Sec", "Page Faults", 1],
    "Transfers/Sec" => ["Total IO tps", "Total I/O Transfers/Sec", "Transfers", 1],
    "Total Used Sockets" => ["Socket totsck", "Total Used Sockets", "Sockets", 1],
    "Trans Carrier Errors/Sec" => ["Net Err DEVICE txcarrps", "DEVICE Transmit Carrier Errors/Sec", "Errors", 1],
    "Trans Errors/Sec" => ["Net Err DEVICE txerrps", "DEVICE Transmit Errors/Sec", "Errors", 1],
    "Trans FIFO Overrun Errors/Sec" => ["Net Err DEVICE txfifops", "DEVICE Transmit FIFO Overrun Errors/Sec", "Errors", 1],
    "Trans Pkts Dropped/Sec" => ["Net Err DEVICE txdropps", "DEVICE Transmit Packets Dropped/Sec", "Packets", 1],
    "UDP D'grams Recv'd/Sec" => ["UDPv4 idgmps", "UDP Datagrams Received/Sec", "UDP Datagrams", 1],
    "UDP D'grams Trans'd/Sec No Port" => ["UDPv4 noportps", "UDP Datagrams Transmitted/Sec No Port", "UDP Datagrams", 1],
    "UDP D'grams Trans'd/Sec Undelivered" => ["UDPv4 idgmerrps", "UDP Datagrams Transmitted/Sec Undelivered", "UDP Datagrams", 1],
    "UDP D'grams Trans'd/Sec" => ["UDPv4 odgmps", "UDP Datagrams Transmitted/Sec", "UDP Datagrams", 1],
    "UDP Packets Recv'd/Sec" => ["NFS Server udpps", "UDP Packets Received'd/Sec", "UDP Packets", 1],
    "UDP Sockets in Use" => ["Socket udpsck", "UDP Sockets in Use", "Sockets", 1],
    "Unused Directory Cache Entries" => ["Files dentunusd", "Unused Directory Cache Entries", "Directory Cache Entries", 1],
    "Used File Handles" => ["Files file-sz", "Used File Handles", "File Handles", 1],
    "Used Inode Handlers" => ["Files inode-nr", "Used Inode Handlers", "Inode Handlers", 1],
    "Used Memory KB" => ["Memory kbmemused", "Used Memory (MBytes)", "MBytes", 1024],
    "Used Swap Space KB" => ["Memory kbswpused", "Used Swap Space (MBytes)", "MBytes", 1024],
    "Write Transfers/Sec" => ["Total IO wtps", "Total I/O Write Transfers/Sec", "Transfers", 1],
);
#
# Read the command-line parameters
#
getopts('D:');

our ($opt_D);
my ($Directory, @XlsFileList, $XlsFile, $IniFile);

if (defined $opt_D) {
    $Directory = $opt_D;
} else {
    $Directory = ".";
}
#
# Open the current working directory, or the one specified with the '-D' switch, and get a list of .xls files therein
#
opendir (DIR, "$Directory$Slash") or die "Can't open directory \"$Directory$Slash\"";
my @AllFileList = readdir(DIR);
closedir (DIR);
for my $File (@AllFileList) {
    push @XlsFileList, $File if $File =~ /\w+\.xls$/i;
}
if (@XlsFileList == 0){
    if (defined $opt_D){
        print "\nNo .xls files have been found in $opt_D\n";
    } else {
        print "\nNo .xls files have been found in your current directory\n";
    }
    print "Please specify a directory containing some .xls files created by sar2xls.pl\n";
    exit;
}
#
# Now prompt the user which sar2xls spreadsheet file to open
#
print "\nThe following spreadsheets have been found in that directory:\n\n";
my $XlsFileIndex = prompt('m',
                        {prompt                    => "Select a spreadsheet to read parameters from > ?",
                        items                      => [ @XlsFileList ],
                        display_base               => 1,
                        return_base                => 0,
                        accept_multiple_selections => FALSE,
                        accept_empty_selection     => FALSE,
                        },'','');
$XlsFile = $XlsFileList[$XlsFileIndex];
#
# Now prompt the user for an output ini file name
#
my $YesNo;
my $Happy = FALSE;
while (!$Happy){
    $IniFile = prompt('x', "\nWhat will the GenGraphs ini file be called > ?", '', 'graphs.ini' );
    $Happy = TRUE;
    for my $File (@AllFileList) {
        if ($File eq $IniFile) {
            $YesNo = &prompt("y", "\n$IniFile already exists, OK to overwrite?", "y/n", "n");
            if (!$YesNo){
                $Happy = FALSE;
            }
        }
    }
}
my ($NewSheetTitle, $GraphTitle, $YaxisTitle, $CellDivisionFactor);
print "\nAre you intending to compare:\n\n";
my $OverlayType = prompt('m',
                        {prompt                    => "Select graph type...",
                        items                      => [ "Multiple systems on the same day",
                                                        "The same system on multiple days",
                                                        "Multiple systems(FQDN) on the same day"],
                        display_base               => 1,
                        return_base                => 0,
                        accept_multiple_selections => FALSE,
                        accept_empty_selection     => FALSE,
                        }, '', '1');    
if ($OverlayType eq 0) {
    $OverlayType = "\$SYSNAME";
} elsif ($OverlayType eq 1) {
    $OverlayType = "\$SARDATE";
} elsif ($OverlayType eq 2) {
    $OverlayType = "\$SYSFQDN";
}

my $Charts = "";
my $Titles = "";
my $Data = "";
my $ChartCount = 0;
print "\nProcessing $XlsFile\n\n";
my $oBook = Spreadsheet::ParseExcel::Workbook->Parse("$Directory$Slash$XlsFile");

for(my $iSheet=0; $iSheet < $oBook->{SheetCount}; $iSheet++) {
    my $Worksheet = $oBook->{Worksheet}[$iSheet];
    my $SheetNamePrinted = 0;
    next if ($Worksheet->{Name} eq "Overview");
    $YesNo = &prompt("y", "Graph anything from the $Worksheet->{Name} page?", "y/n", "n");
    if ($YesNo) {
        #
        # Read the headings
        #
        for (my $Column = 1; defined $Worksheet->{MaxCol} && $Column <= $Worksheet->{MaxCol}; $Column++) {
            my $Cell = $Worksheet->{Cells}[0][$Column];
            my $Heading = $Cell->{Val};
            $YesNo = &prompt("y", "\tGraph $Heading?", "y/n", "n");
            if ($YesNo) {
                $ChartCount++;
                $NewSheetTitle = $GraphAttributes{$Heading}[0];
                $GraphTitle = $GraphAttributes{$Heading}[1];
                $YaxisTitle = $GraphAttributes{$Heading}[2];
                $CellDivisionFactor = $GraphAttributes{$Heading}[3];
        if (($Worksheet->{Name} =~ m/^IO - (.+)$/) || ($Worksheet->{Name} =~ m/^Network - (.+)$/) || ($Worksheet->{Name} =~ m/^Network Errors - (.+)$/)) {
            my $Device = $1;
            my $PrettyDevice = $Device;
            #
            # Swap back any '~' characters that a device map may have put in for the true '/'
            #
            $PrettyDevice =~ s/~/\//;
            if (($Device =~ /^mapper/) || ($Device =~ /^cciss/)){
                $PrettyDevice = "/dev/" . $PrettyDevice;
            } elsif ($Device =~ /^dev/){
                $PrettyDevice = "/" . $PrettyDevice;
            } 
            $GraphTitle =~ s/DEVICE/$PrettyDevice/;
            #
            # Calculate the length of the New Sheet Title string when we substitute the device name into it
            #
            my $NewSheetTitleLength = length($NewSheetTitle) - length("DEVICE") + length($Device);
            if ($NewSheetTitleLength > MAX_SHEET_NAME_LENGTH){
                #
                # we have to trim some off the device name to fit into out maximum sheet title length
                #
                my $TrimLength = $NewSheetTitleLength - MAX_SHEET_NAME_LENGTH + length("...");
                $Device = "..." . substr($Device, $TrimLength);
            }
            $NewSheetTitle =~ s/DEVICE/$Device/;
        }
        if ($Worksheet->{Name} =~ m/^CPU - (.+$)/) {
            my $CPU = $1;
            $NewSheetTitle =~ s/CPU/CPU $CPU/;
            $GraphTitle =~ s/CPU/CPU:$CPU/;
        }
                $Charts .= "Chart$ChartCount=$NewSheetTitle\n";
                $Titles .= "\n[Chart${ChartCount}Titles]\nGraphTitle=$GraphTitle\nYAxisTitle=$YaxisTitle\n";
                if ($CellDivisionFactor == 1){
                    $Data .= "\n[Chart${ChartCount}Data]\n$GraphTitle - $OverlayType=$Worksheet->{Name}::$Cell->{Val}\n";
                } else {
                    $Data .= "\n[Chart${ChartCount}Data]\n$GraphTitle - ${OverlayType}::$CellDivisionFactor=$Worksheet->{Name}::$Cell->{Val}\n";
                }
            }
        }
    }
}
#
# Now prompt for the Save file name containing the charts
#
$Happy = FALSE;
my $SaveFileName;
while (!$Happy){
    $SaveFileName = prompt('x', "\nWhat will the GenGraphs output spreadsheet be called > ?", '', 'graphs.xlsx' );
    $Happy = TRUE;
    for my $File (@AllFileList) {
        if ($File eq $SaveFileName) {
            $YesNo = &prompt("y", "\n$SaveFileName already exists, OK to overwrite?", "y/n", "n");
            if (!$YesNo){
                $Happy = FALSE;
            }
        }
    }
}
my $SaveFileNameAbsPath = File::Spec->rel2abs("$Directory$Slash$SaveFileName");
#
# Open the output file and write the text
#
open (INIFILE, ">$Directory$Slash$IniFile") or die "Can't Create $Directory$Slash$IniFile \n";
print INIFILE "[General]\n";
print INIFILE "SaveFileName=$SaveFileNameAbsPath\n\n";
print INIFILE "[Files]\n";
print INIFILE ";\n";
print INIFILE "; List the files that we want to get the source data from\n";
print INIFILE ";\n";
#
# Prompt for the .xls files to be included as the source sheets
#
print "\nSelect the spreadsheets to be used as source files for the GenGraphs plots:\n\n";
my @SourceXlsFileIndices = prompt('m',
                        {prompt                    => "",
                        items                      => [ @XlsFileList ],
                        display_base               => 1,
                        return_base                => 0,
                        accept_multiple_selections => TRUE,
                        accept_empty_selection     => TRUE,
                        ignore_whitespace          => FALSE,
                        separator                  => ' '
                        },'','');

my $FileCount = 0;
for my $SourceXlsFileIndex (@SourceXlsFileIndices) {
    $FileCount++;
    $XlsFile = $XlsFileList[$SourceXlsFileIndex];
    my $XlsFileAbsPath = File::Spec->rel2abs("$Directory$Slash$XlsFile");
    print INIFILE "File$FileCount=$XlsFileAbsPath\n";
}
print INIFILE ";\n";
print INIFILE "; List the charts that we want to create - these names will form the sheet titles\n";
print INIFILE ";\n";
print INIFILE "[Charts]\n";
print INIFILE $Charts;
print INIFILE $Titles;
print INIFILE $Data;
if ($OSNAME eq "MSWin32") {
    $YesNo = &prompt("y", "\nLaunch GenGraphs.py using the created ini file > ?", "y/n", "n");
    if ($YesNo) {
        if (defined $opt_D) {
            system("python", "GenGraphs.py", "-f", "$IniFile", "-D", "$Directory");
        } else {
            system("python", "GenGraphs.py", "-f", "$IniFile");
        }
    }
}
exit(1);

#-------------------------------------------------------------------------------
# Subroutine:   PrintUsage
# Function:     Prints a usage message to the console
# Arguments:    None
# Returns:      Nothing
#-------------------------------------------------------------------------------
sub PrintUsage {
    print "\nusage: GenIni.pl [-D directory]\n\n";
    print "Where:\n";
    print "    -D Specifies an optional directory to use for input and output\n"; 
    exit;
}
