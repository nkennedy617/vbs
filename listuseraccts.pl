use Win32;
use Win32::AdminMisc;
use win32::perms; 

if( Win32::AdminMisc::GetUsers( "\\hogwarts.cs.clemson.edu", "", \@List ) )
{
    $USER = "c:/ME/NEW/user.txt";
    open USER, ">$USER" or die "ERRROR\n";
    map { print USER "\t$_\n";} @List;
    close (USER);
}
else{
	print "error\n";
}