#!C:/Perl/bin/perl -w
use strict;
use Cwd;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use IO::File;
use File::Slurp;


$Win32::OLE::Warn = 3; # Die on Errors.

# ::Warn = 2; throws the errors, but #
# expects that the programmer deals  #

### Win 32 seems to expect the entire path of the excel file
my $dir = getcwd;
my $excelfile = shift @ARGV;
$excelfile = $dir.'\\'.$excelfile;

my $files=shift @ARGV;

#First, we need an excel object to work with, so if there isn't an open one, we create a new one, and we define how the object is going to exit

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
        || Win32::OLE->new('Excel.Application', 'Quit');

#For the sake of this program, we'll turn off all those pesky alert boxes, such as the SaveAs response "This file already exists", etc. using the DisplayAlerts property.

$Excel->{DisplayAlerts}=0;   

#opened an existing file to work with 
                                                 
my $Book = $Excel->Workbooks->Open($excelfile);   

#Now we create a reference to a worksheet object and activate the sheet to give it focus so that actions taken on the workbook or application objects occur on this sheet unless otherwise specified.

 my $Sheet = $Book->Worksheets("Sheet1");
$Sheet->Activate();  

my $usedRange = $Sheet->UsedRange()->{Value};

main();

sub main 
{
	##read rows; create METS file for each

	foreach my $row (@$usedRange) 
	{
		
		#my $outputfile = "Brooker-".sprintf("%04d", @$row[1])."mets.xml";
		my $outputfile = "Brooker-".@$row[1]."mets.xml";

		my $fh = IO::File->new($outputfile, 'w')
			or die "unable to open output file for writing: $!";

		metsHdr($fh, $row);	
		mods($fh, $row);
		fileSec($fh, $row);
		structMap($fh, $row);
		closeMets($fh);
   		$fh->close();
	}
}

#############
sub structMap
#############
{
	my $fh=shift;
	my $row=shift;
	my ($year,$number,$type,$primaryLocation,$description,$names,$otherLocations,$labels) = @$row;
	#$number = sprintf("%04d", $number);
	$fh->print("\t<mets:structMap TYPE=\"physical\" ID=\"SMD1\">\n");
	$fh->print("\t\t<mets:div TYPE=\"manuscript\" LABEL=\"The Robert E. Brooker III Collection of American Legal and Land Use Documents. No. $number.\" ORDER=\"1\" DMDID=\"DMD1\">\n");

	###Second level divs
	my @labelList;

	if ($labels)
	{
		
		@labelList = split(";", $labels);

	}
	else 
	{
		@labelList = ("Side a", "Side b", "Side c", "Side d", "Side e", "Side f", "Side g", "Side h", "Side i", "Side j", "Side k", "Side l", "Side m", "Side n", "Side o", "Side p", "Side q", "Side r", "Side s", "Side t", "Side u", "Side v", "Side w", "Side x", "Side y", "Side z","Side aa", "Side ab", "Side ac", "Side ad", "Side ae", "Side af", "Side ag", "Side ah", "Side ai", "Side aj", "Side ak", "Side al", "Side am", "Side an", "Side ao", "Side ap", "Side aq", "Side ar", "Side as", "Side at", "Side au", "Side av", "Side aw", "Side ax", "Side ay", "Side az", );
	}
	
		my $i=0;
	my @file = read_file($files);
	foreach(@file)
	{
		chomp;
		if ((substr($_ , 8 , 4) eq $number) || (substr($_ , 8 , 7) eq $number))
		{
			$i++;
			$labelList[$i-1] =~ s/^\s+|\s+$//g;
			$labelList[$i-1] =~ s/&/&amp;/g;
			$fh->print("\t\t\t<mets:div TYPE=\"page\" LABEL=\"No. $number\/". $labelList[$i-1] ."\" ORDER=\"$i\" ORDERLABEL=\"".lc($labelList[$i-1])."\">\n");
	#####fptrs
			$fh->print("\t\t\t\t<mets:fptr FILEID=\"t".sprintf("%04d", $i)."\"/>\n"); 
			$fh->print("\t\t\t\t<mets:fptr FILEID=\"jp".sprintf("%04d", $i)."\"/>\n"); 
			$fh->print("\t\t\t\t<mets:fptr FILEID=\"j2k".sprintf("%04d", $i)."\"/>\n"); 
			$fh->print("\t\t\t</mets:div>\n");
		}
	}
	#####close up structMap
	$fh->print("\t\t</mets:div>\n");
	$fh->print("\t</mets:structMap>\n");
}

###########
sub fileSec
###########
{
	my $fh=shift;
	my $row=shift;
	my ($year,$number,$type,$primaryLocation,$description,$names,$otherLocations) = @$row;
	#$number = sprintf("%04d", $number);

	my @file = read_file($files);

	$fh->print("\t<mets:fileSec ID=\"FSD1\">\n");


	####archive fileGrp
	$fh->print("\t\t<mets:fileGrp USE=\"archive\">\n");

	my $i=0;
	foreach(@file)
	{
		chomp;
		if ((substr($_ , 8 , 4) eq $number) || (substr($_ , 8 , 7) eq $number))
	

		{
			$i++;
			#die "WARNING: more than 52 componenent files\n"
			#	if ($i eq 53);  #script only set up to handle 52 components
						    #add more components if script fails	
			$fh->print("\t\t\t<mets:file ID=\"t" . sprintf("%04d", $i) ."\" MIMETYPE=\"image\/tiff\" GROUPID=\"GID".$i."\" SEQ=\"$i\">\n");
			$fh->print("\t\t\t\t<mets:FLocat xlink:href=\"file://streams/" . $_ . ".tif\" LOCTYPE=\"URL\"\/>\n"); 
			$fh->print("\t\t\t<\/mets:file>\n");
		}
	}
	$fh->print("\t\t<\/mets:fileGrp>\n");
	####jpg fileGrp

	$fh->print("\t\t<mets:fileGrp USE=\"reference image\">\n");

	$i=0;
	foreach(@file)
	{
		chomp;
		if ((substr($_ , 8 , 4) eq $number) || (substr($_ , 8 , 7) eq $number))
		{
			$i++;
			$fh->print("\t\t\t<mets:file ID=\"jp".sprintf("%04d", $i)  ."\" MIMETYPE=\"image\/jpeg\" GROUPID=\"GID".$i."\" SEQ=\"$i\">\n");
			$fh->print("\t\t\t\t<mets:FLocat xlink:href=\"file://streams/" . $_ . ".jpg\" LOCTYPE=\"URL\"\/>\n");
			$fh->print("\t\t\t<\/mets:file>\n");
		}
	}
	$fh->print("\t\t<\/mets:fileGrp>\n");

	####jpg 2000 fileGr[

	$fh->print("\t\t<mets:fileGrp USE=\"reference image dynamic\">\n");
	$i=0;
	foreach(@file)
	{
		chomp;

		if ((substr($_ , 8 , 4) eq $number) || (substr($_ , 8 , 7) eq $number))
		{
			$i++;
			$fh->print("\t\t\t<mets:file ID=\"j2k".sprintf("%04d", $i)."\" MIMETYPE=\"image\/jp2\" GROUPID=\"GID".$i."\" SEQ=\"$i\">\n");
$fh->print("\t\t\t\t<mets:FLocat xlink:href=\"file://streams/" . $_ . ".jp2\" LOCTYPE=\"URL\"\/>\n");
			$fh->print("\t\t\t<\/mets:file>\n");

		}
	}
	$fh->print("\t\t<\/mets:fileGrp>\n");

	####Close up fileSec
	$fh->print("\t</mets:fileSec>\n");
}

########
sub mods
########
{
	my $fh=shift;
	my $row=shift;
	my ($year,$number,$type,$primaryLocation,$description,$names,$otherLocations) = @$row;
	my $jr;

	#$number = sprintf("%04d", $number);

	print "\n\n$number, $type, $primaryLocation\n";

	$fh->print("<mets:dmdSec ID=\"DMD1\">\n");
	$fh->print("\t\t<mets:mdWrap MDTYPE=\"MODS\">\n");
	$fh->print("\t\t\t<mets:xmlData>\n");

	$fh->print("\t\t\t\t<mods:mods>\n");

	$fh->print("\t\t\t\t\t<mods:titleInfo>\n");
	$fh->print("\t\t\t\t\t\t<mods:nonSort>The <\/mods:nonSort>\n");
	$fh->print("\t\t\t\t\t\t<mods:title>Robert E. Brooker III Collection of American Legal and Land Use Documents, 1716-1930<\/mods:title>\n");
	$fh->print("\t\t\t\t\t\t<mods:partNumber>No. $number<\/mods:partNumber>\n");
	$fh->print("\t\t\t\t\t<\/mods:titleInfo>\n");

	my @names = split(/\s*,\s*/, $names);
	foreach (@names) 
	{
		$_=~s/\?//g;
		if ($_ =~ m/Jr./) {$jr="true"; $_ =~ s/Jr.//;} else {$jr="false";}
		if (m/(\w+\'*\w*\-*\w*\'*\w*)\s*$/)  {#oct20
		my $family_name=$1;

		my $given_name=$`;
		$given_name =~ s/\s*$//;

		$fh->print("\t\t\t\t\t<mods:name type=\"personal\">\n");
		$fh->print("\t\t\t\t\t\t<mods:namePart type=\"family\">$family_name<\/mods:namePart>\n");
		$fh->print("\t\t\t\t\t\t<mods:namePart type=\"given\">$given_name<\/mods:namePart>\n");
		if ($jr eq 'true') {$fh->print("\t\t\t\t\t\t<mods:namePart type=\"termsOfAddress\">Jr.<\/mods:namePart>\n");}
		if ($jr eq 'true'){$fh->print("\t\t\t\t\t\t<mods:displayForm>$family_name, $given_name, Jr.<\/mods:displayForm>\n");}
		else 
		{
			$fh->print("\t\t\t\t\t\t<mods:displayForm>$family_name, $given_name<\/mods:displayForm>\n");}
			$fh->print("\t\t\t\t\t<\/mods:name>\n");
		}

		else 
		{
			$fh->print("\t\t\t\t\t<mods:name type=\"personal\">\n");
			$fh->print("\t\t\t\t\t\t<mods:namePart type=\"family\">$_<\/mods:namePart>\n");
			$fh->print("\t\t\t\t\t\t<mods:namePart type=\"given\"><\/mods:namePart>\n");
			$fh->print("\t\t\t\t\t\t<mods:displayForm>$_<\/mods:displayForm>\n");
			$fh->print("\t\t\t\t\t<\/mods:name>\n"); 
		};
	}


	$fh->print("\t\t\t\t\t<mods:typeOfResource manuscript=\"yes\">text<\/mods:typeOfResource>\n");

	my @type = split(/\s*;\s*/, $type);
	foreach (@type) 
	{  
		s/^\s+//;
		s/\s+$//;
		$_=lc($_);
		$fh->print("\t\t\t\t\t<mods:genre authority=\"local\" type=\"workType\">$_<\/mods:genre>\n");
	}

	$fh->print("\t\t\t\t\t<mods:originInfo>\n");

	if ($year =~ m/\d\d\d\d/)
	{
		$fh->print("\t\t\t\t\t\t<mods:dateCreated>$year<\/mods:dateCreated>\n");
		$fh->print("\t\t\t\t\t\t<mods:dateCreated encoding=\"w3cdtf\" keyDate=\"yes\">$year<\/mods:dateCreated>\n");
	}
	elsif ($year =~ m/unknown/i)
	{
		$fh->print("\t\t\t\t\t\t<mods:dateCreated>$year<\/mods:dateCreated>\n");
		$fh->print("\t\t\t\t\t\t<mods:dateCreated encoding=\"w3cdtf\" point=\"start\" keyDate=\"yes\">1716<\/mods:dateCreated>\n");
		$fh->print("\t\t\t\t\t\t<mods:dateCreated encoding=\"w3cdtf\" point=\"end\">1930<\/mods:dateCreated>\n");


	}

	$fh->print("\t\t\t\t\t\t<mods:issuance>monographic<\/mods:issuance>\n");
	$fh->print("\t\t\t\t\t<\/mods:originInfo>\n");
	$fh->print("\t\t\t\t\t<mods:language>\n");
	$fh->print("\t\t\t\t\t\t<mods:languageTerm type=\"text\">English<\/mods:languageTerm>\n");
	$fh->print("\t\t\t\t\t\t<mods:languageTerm authority=\"iso639-2b\" type=\"code\">eng<\/mods:languageTerm>\n");
	$fh->print("\t\t\t\t\t<\/mods:language>\n");

	$fh->print("\t\t\t\t\t<mods:physicalDescription>\n");
	$fh->print("\t\t\t\t\t\t<mods:form authority=\"marcform\">electronic<\/mods:form>\n");
	$fh->print("\t\t\t\t\t\t<mods:internetMediaType>image\/jpeg<\/mods:internetMediaType>\n");
	$fh->print("\t\t\t\t\t\t<mods:internetMediaType>image\/jp2<\/mods:internetMediaType>\n");
	$fh->print("\t\t\t\t\t\t<mods:internetMediaType>image\/tiff<\/mods:internetMediaType>\n");
	$fh->print("\t\t\t\t\t\t<mods:digitalOrigin>reformatted digital<\/mods:digitalOrigin>\n");

	if ($description && ($description =~ m/facsim/)) 
	{
		$fh->print("\t\t\t\t\t\t<mods:extent>1 color facsimile of original manuscript (front side only)</mods:extent>\n");
	} 
	else {$fh->print("\t\t\t\t\t\t<mods:extent>1 manuscript</mods:extent>\n");};

	$fh->print("\t\t\t\t\t<\/mods:physicalDescription>\n");


	if ($description) 
	{
		my $abstract = $description;
		$abstract =~s/ Color facsimile copy of original document\.//; 
		$abstract =~ s/&/&amp;/g;
		$fh->print("\t\t\t\t\t<mods:abstract>$abstract<\/mods:abstract>\n");
	}

	my $namesList = $names;
	$namesList =~ s/\?/\[\?\]/g;
	$fh->print("\t\t\t\t\t<mods:note>Names in document: $namesList.<\/mods:note>\n");
	$fh->print("\t\t\t\t\t<mods:note>Primary location: $primaryLocation.<\/mods:note>\n");

	if ($otherLocations ne "None") {$fh->print("\t\t\t\t\t<mods:note>Other locations: $otherLocations.<\/mods:note>\n");};
	$fh->print("\t\t\t\t\t<mods:note type=\"reproduction\">Electronic reproduction. Chestnut Hill, Mass. : University Libraries, Boston College, 2016.<\/mods:note>\n");
	$fh->print("\t\t\t\t\t<mods:note type=\"original location\">Brooker Collection, Daniel R. Coquillette Rare Book Room, Boston College Law Library.<\/mods:note>\n");

	if ($primaryLocation ne "Unknown") {modsHierarchicalGeographic($primaryLocation, $fh)};

	if ($otherLocations ne "None") {modsHierarchicalGeographicOther($otherLocations, $fh);};

	$fh->print("\t\t\t\t\t<mods:identifier type=\"local\">$number<\/mods:identifier>\n");

	$fh->print("\t\t\t\t\t<mods:accessCondition type=\"useAndReproduction\">Use of this resource is governed by the terms and conditions of the \"Creative Commons Attribution-Noncommercial 3.0 United States\" license \(http://creativecommons.org/licenses/by-nc/3.0/us/\)<\/mods:accessCondition>\n");

	$fh->print("\t\t\t\t\t<mods:extension>\n");
	$fh->print("\t\t\t\t\t\t<localCollectionName>Brooker<\/localCollectionName>\n");
	$fh->print("\t\t\t\t\t<\/mods:extension>\n");

	$fh->print("\t\t\t\t\t<mods:recordInfo>\n");
	$fh->print("\t\t\t\t\t\t<mods:recordContentSource>MChB<\/mods:recordContentSource>\n");
	$fh->print("\t\t\t\t\t\t<mods:recordIdentifier source=\"Brooker\">$number<\/mods:recordIdentifier>\n");
	$fh->print("\t\t\t\t\t\t<mods:languageOfCataloging>\n");
	$fh->print("\t\t\t\t\t\t\t<mods:languageTerm type=\"text\">English<\/mods:languageTerm>\n");
	$fh->print("\t\t\t\t\t\t\t<mods:languageTerm type=\"code\" authority=\"iso639-2b\">eng<\/mods:languageTerm>\n");

	$fh->print("\t\t\t\t\t\t<\/mods:languageOfCataloging>\n");
	$fh->print("\t\t\t\t\t<\/mods:recordInfo>\n");
		
	$fh->print("\t\t\t\t<\/mods:mods>\n");
	$fh->print("\t\t\t<\/mets:xmlData>\n");
	$fh->print("\t\t<\/mets:mdWrap>\n");
	$fh->print("\t<\/mets:dmdSec>\n");


}

##############################
sub modsHierarchicalGeographic 
##############################
{
	my $location = shift;
	my $fh = shift;

	$location =~ m/(^[A-Za-z\s]*)/;
	my $state = $1;
	my $townCounty = $';
	$state =~ s/\s+$//;

	if ($state)
	{

		$fh->print("\t\t\t\t\t<mods:subject>\n");
		$fh->print("\t\t\t\t\t\t<mods:hierarchicalGeographic>\n");
		$fh->print("\t\t\t\t\t\t\t<mods:state>$state<\/mods:state>\n");

	}	

	#town and county
	if ($townCounty =~ m/\(([A-Za-z\s]*), ([A-Za-z\s*]*)\){1}/)
	{	
		my $town = $1;
		my $county = $2;
		$county =~ s/in //;

		print " town is: $town and county is $county\n";
		if ($county) {$fh->print("\t\t\t\t\t\t\t<mods:county>$county<\/mods:county>\n");};
		if ($town) {$fh->print("\t\t\t\t\t\t\t<mods:city>$town<\/mods:city>\n");};
		print "$location\n";
	}

	#town only
	if ($townCounty =~ m/\(([A-Za-z\s]*)\){1}$/)
	{	
		my $town = $1;
		if ($town) {$fh->print("\t\t\t\t\t\t\t<mods:city>$town<\/mods:city>\n");};
		print "$location\n";
	}

	elsif ($townCounty =~ m/\(([A-Za-z\s]*) \(([A-Za-z\s*]*)\)\)/)
	{
		my $town = $1;
		my $citySection=$2;
		if ($town) {$fh->print("\t\t\t\t\t\t\t<mods:city>$town<\/mods:city>\n");};
		if ($town) {$fh->print("\t\t\t\t\t\t\t<mods:citySection>$citySection<\/mods:citySection>\n");};
			
	};

	$fh->print("\t\t\t\t\t\t<\/mods:hierarchicalGeographic>\n");
	$fh->print("\t\t\t\t\t<\/mods:subject>\n");

};

###################################
sub modsHierarchicalGeographicOther 
###################################
{
	my $otherLocations = shift;
	my $fh = shift;

	my @otherLocations = split (/; /, $otherLocations);

	foreach (@otherLocations) 
	{
		$fh->print("\t\t\t\t\t<mods:subject>\n");
		$fh->print("\t\t\t\t\t\t<mods:hierarchicalGeographic>\n");
	 	print "other location: $_\n";
		print "$_ is dollar under\n";
		my @otherLocationComponents = split (/, /, $_);

		if (scalar(@otherLocationComponents == 1)) 
		{
			$fh->print("\t\t\t\t\t\t\t<mods:state>$otherLocationComponents[0]<\/mods:state>\n");
		}

		elsif (scalar(@otherLocationComponents == 2)) 
		{
			$fh->print("\t\t\t\t\t\t\t<mods:state>$otherLocationComponents[1]<\/mods:state>\n");
			$fh->print("\t\t\t\t\t\t\t<mods:city>$otherLocationComponents[0]<\/mods:city>\n");
		}

		elsif (scalar(@otherLocationComponents == 3)) 
		{
			$fh->print("\t\t\t\t\t\t\t<mods:state>$otherLocationComponents[2]<\/mods:state>\n");
			$fh->print("\t\t\t\t\t\t\t<mods:county>$otherLocationComponents[1]<\/mods:county>\n");
			$fh->print("\t\t\t\t\t\t\t<mods:city>$otherLocationComponents[0]<\/mods:city>\n");
		};

		$fh->print("\t\t\t\t\t\t<\/mods:hierarchicalGeographic>\n");
		$fh->print("\t\t\t\t\t<\/mods:subject>\n");
	}
}

###########
sub metsHdr
###########
{
	my $fh=shift;
	my $row=shift;
	my ($year,$number,$type,$primaryLocation,$description,$names,$otherLocations) = @$row;
	#$number = sprintf("%04d", $number);

	$fh->print("<?xml version='1.0' encoding='UTF-8' ?>\n");
	$fh->print("<mets:mets OBJID=\"law.brooker.$number\" LABEL=\"The Robert E. Brooker III Collection of American Legal and Land Use Documents. No. $number\" TYPE=\"manuscript\" xmlns:mets=\"http://www.loc.gov/METS/\"
    xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"
    xmlns:mods=\"http://www.loc.gov/mods/v3\"
    xsi:schemaLocation=\"http://www.loc.gov/METS/ http://www.loc.gov/standards/mets/mets.xsd http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-3.xsd\">
\n");
	my ($sec,$min,$hour,$mday,$mon,$yr,$wday,$yday,$isdst)=localtime();
$fh->print("<mets:metsHdr CREATEDATE=\"".($yr+1900)."-".sprintf("%02d",$mon+1)."-".sprintf("%02d",$mday)."T".sprintf("%02d",$hour).":".sprintf("%02d",$min).":".sprintf("%02d",$min)."\">\n");
	$fh->print("\t<mets:agent ROLE=\"CREATOR\" TYPE=\"ORGANIZATION\">\n");
	$fh->print("\t\t<mets:name>BC Digital Library Program<\/mets:name>\n");
	$fh->print("\t<\/mets:agent>\n");
	$fh->print("<\/mets:metsHdr>\n");
}

#############
sub closeMets
#############
{
	my $fh=shift;
	$fh->print("<\/mets:mets>\n");
}

=pod
Usage: brooker.pl metadata.xlsx componentslist.txt 

metadata.xlsx:  A metadata file generated by the Law Library.  This file is enhanced post-scanning with labels for any documents that have complicated labelling (i.e. other than side a; side b; side c ..... The script expects the sheet containing the metadata to be named Sheet1

componentslist.txt is a list of file prefixes used for the batch

Betsy Post, Boston College. 
betsy.post@bc.edu

=cut
