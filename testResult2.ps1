#
# testResult2.ps1	by Amina JEBARI
#

  $testFolder_ = "C:\Users\Amina\Desktop\test\results"
  $xsltTemplate_ ="C:\Users\Amina\Desktop\test\testResult.xsl" 
  $endFolder_ = "C:\Users\Amina\Desktop\final"
  $xmltohtml_ = "C:\Users\Amina\Desktop\test\convertXmlToHtml" 
  $outputhtmlfile_ =  "C:\Users\Amina\Desktop\final\aminatest.html"
  $xmlfile_ =  "C:\Users\Amina\Desktop\final\ranorex.xml"
  $xslttemplatefinal_ =  "C:\Users\Amina\Desktop\final\testResult.xsl"



function replace-string([string] $file, [string] $old, [string] $new)
{
    (get-content $file) | foreach-object {$_ -replace [regex]::Escape($old), $new} | set-content  $file
}

function CopyItem($source_, $destination_, $step_, $errorcode_, $outputhtmlfile_)
{
    if(!(Test-Path -Path $destination_ )){
        New-Item -ItemType directory -Path $destination_ | Out-Null
    }
    $copystatus = Copy-Item $source_ -Destination $destination_ -Recurse -Force -PassThru -ErrorAction silentlyContinue
    if ($copystatus) {
        Write-Verbose "$step_ completed successfully." 
    } else { 
        Write-Error "$step_ failed." 
        Exit $errorcode_
    }
}

function main($testFolder_, $xsltTemplate_, $endFolder_, $xmltohtml_,$outputhtmlfile_) 
{
	# variable
	$cpt = 0
	$Variable = ""
	$allVariable = ""
	$leftPanel = ""
	$allLeftPanel = ""
	$body = ""
	$allBody = ""
	

	# first we copy the xslt template to our end folder
	copyItem $xsltTemplate_ $endFolder_ "Copying the xslt template in the end folder" 2

	# then we get it in the variable $xsltfile to write in it
	$xsltName = "testResult.xsl"
	[xml]$xsltfile = Get-Content "$endFolder_\$xsltName"

	# we look how many test results in xml files there are 
	$files = Get-Item "$testFolder_\*"


	# for each of them we are going to write the informations into the xslt template
	Foreach ($file in $files)
	{
		[xml]$ConfigFile = Get-Content $file
		$cpt+=1
		$fileName = $file.Name

		if ($file.Name -match "vs")
		{
			
			$name = $fileName.replace(".xml","" )
			$Variable = "<xsl:variable name=`"name$cpt`" select=`"'$name'`" /><xsl:variable name=`"path$cpt`" select=`"document('$fileName')/TestRun`" /><xsl:variable name=`"error$cpt`" select=`"document('$fileName')/TestRun/ResultSummary/Counters/@failed`" />"
			$body = "<xsl:call-template name=`"visual-studio`"><xsl:with-param name=`"path`" select=`"`$path$cpt`"/><xsl:with-param name=`"name`" select=`"`$name$cpt`"/></xsl:call-template>"
		}
		else 
		{	
			$name = $fileName.replace(".xml","" )
			$Variable = "<xsl:variable name=`"name$cpt`" select=`"'$name'`" /><xsl:variable name=`"path$cpt`" select=`"document('$fileName')/report`" /><xsl:variable name=`"error$cpt`" select=`"document('$fileName')/report/activity/activity/@totalerrorcount`" />"
			$body = "<xsl:call-template name=`"ranorex`"><xsl:with-param name=`"path`" select=`"`$path$cpt`"/><xsl:with-param name=`"name`" select=`"`$name$cpt`"/></xsl:call-template>"
		}
		$leftPanel = "<xsl:call-template name=`"left-panel`"><xsl:with-param name=`"name`" select=`"`$name$cpt`"/><xsl:with-param name=`"path`" select=`"`$error$cpt`"/></xsl:call-template>"

		# put all the parts together
		$allVariable += $Variable
		$allLeftPanel += $leftPanel
		$allBody += $body

		# copying each xml into the end folder
		copyItem "$testFolder_\$fileName" $endFolder_ "Copying the xslt template in the end folder" 2

		# copying each xml into the end folder witch a new encoding
		Get-Content  "$endFolder_\$fileName" | Out-File -Encoding UTF8  "$endFolder_\$fileName+'1'"


		# add the link to the xslt template into each xml file
		replace-string "$endFolder_\$fileName+'1'" `
		"<?xml version=`"1.0`" encoding=`"UTF-8`"?>" `
		"<?xml version=`"1.0`"?>
		<?xml-stylesheet type=`"text/xsl`" href=`"testResult.xsl`"?>"



		$endxml = $file.FullName
	}
	# add the strings to the xslt template
	# 1-variable
	replace-string "$endFolder_\$xsltName" `
	"<!-- set the variables -->" `
	"$allVariable" `
	# 2-left panel
	replace-string "$endFolder_\$xsltName" `
	"<!-- set the left panel -->" `
	"$allLeftPanel" `
	# 3-body
	replace-string "$endFolder_\$xsltName" `
	"<!-- set the body -->" `
	"$allBody" `

        

	<#
	 la conversion du xsl-xml en html
	#>	

	$XsltSettings = New-Object System.Xml.Xsl.XsltSettings($true, $false);
	$XsltSettings.EnableDocumentFunction = 1
	$xslt = New-Object System.Xml.Xsl.XslCompiledTransform
	 $xslt.Load("C:\Users\Amina\Desktop\final\testResult.xsl", $XsltSettings, (New-Object System.Xml.XmlUrlResolver))

	#$XsltSettings = New-Object System.Xml.Xsl.XsltSettings($true, $false)
	
	   
	$xslt.Transform("C:\Users\Amina\Desktop\final\ranorex.xml+'1'", $outputhtmlfile_)

}

<# 
main is called with :
param 1: where the xml results files are 
param 2: where the xslt template is 
param 3: where you want to put the report
param 4: where the xml to html console application is
#>
main $testFolder_ $xsltTemplate_ $endFolder_ $xmltohtml_ $outputhtmlfile_

