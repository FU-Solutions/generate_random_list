<#
.SYNOPSIS
	Will generate a random sample list - for audits or so - from a given list.
	
.DESCRIPTION
	In case you need to generate a sample for your internal control system or audit, you could use this script.
	The script then will generate a new Excelfile with the samples for you or the auditor. Completly random.

.Parameter SourceFile 
	Defines the sourcefile from which the sample is generated.
	No default set - please define an input file.
.Parameter OutFile 
	Defines the outputfile for the sample.
	Default is iks-sample.xlsx.
.Parameter SampleSize
	Defines the amount of entries in your sample-list.
	Default are 40 samples.
.Parameter poWer
    This switch allows you to use more power for your request! Just use switch -w and lay back.
	
.INPUTS
	You have to define an excel-list with parameter SourceFile.
	In general it will get 40 entries from you List - you can change that

.OUTPUTS
	You will get an excel-list with the same data as the inputfile but only the amount of columns you chose.

.EXAMPLE
	Needs to be written.

.NOTES
	File Name		: generate_random_list.ps1
	Author			: Ulli Weichert
	Contact			: ulli@weichert.it - Weichert.IT
	Version			: 0.2
    Release         : Beavers are bored.
	
	Introduction:
		In case you need to generate a sample for your internal control system, you could use this script.
		The script then will generate a new Excelfile with the samples for you. Completly random.
        
	Script details:
		Needs to be written.

.LINK
	https://github.com/w3ich3rt/generate_random_list
.LINK
	<imagine the link to the block here>
#>

## Parameters and variables

param(
	[Parameter(Mandatory=$true,HelpMessage="Sourceinputfile for sample creation.")]
	[Alias('source')]
	[string]$SourceFile,
	[Parameter(Mandatory=$false,HelpMessage="File where the output will be written.")]
	[Alias('output')]
	[string]$OutFile="iks-sample.xlsx",
	[Parameter(Mandatory=$false,HelpMessage="Amount of sample in you generated list")]
	[Alias('samples')]
	[Int]$SampleSize="40",
    [Parameter(Mandatory=$false,HelpMessage="Will provide more power!")]
    [Alias('w')]
    [switch]$poWer=$true
)

## install & import modules or namespaces
Install-Module -Name ImportExcel -scope CurrentUser
Import-Module -Name ImportExcel

# function
function use-power () {
    if ( $poWer ) {
        $wumms = @"
__/\\\______________/\\\__/\\\________/\\\__/\\\________/\\\__/\\\\____________/\\\\__/\\\\____________/\\\\_____/\\\\\\\\\\\_______/\\\____        
 _\/\\\_____________\/\\\_\/\\\_______\/\\\_\/\\\_______\/\\\_\/\\\\\\________/\\\\\\_\/\\\\\\________/\\\\\\___/\\\/////////\\\___/\\\\\\\__       
  _\/\\\_____________\/\\\_\/\\\_______\/\\\_\/\\\_______\/\\\_\/\\\//\\\____/\\\//\\\_\/\\\//\\\____/\\\//\\\__\//\\\______\///___/\\\\\\\\\_      
   _\//\\\____/\\\____/\\\__\/\\\_______\/\\\_\/\\\_______\/\\\_\/\\\\///\\\/\\\/_\/\\\_\/\\\\///\\\/\\\/_\/\\\___\////\\\_________\//\\\\\\\__     
    __\//\\\__/\\\\\__/\\\___\/\\\_______\/\\\_\/\\\_______\/\\\_\/\\\__\///\\\/___\/\\\_\/\\\__\///\\\/___\/\\\______\////\\\_______\//\\\\\___    
     ___\//\\\/\\\/\\\/\\\____\/\\\_______\/\\\_\/\\\_______\/\\\_\/\\\____\///_____\/\\\_\/\\\____\///_____\/\\\_________\////\\\_____\//\\\____   
      ____\//\\\\\\//\\\\\_____\//\\\______/\\\__\//\\\______/\\\__\/\\\_____________\/\\\_\/\\\_____________\/\\\__/\\\______\//\\\_____\///_____  
       _____\//\\\__\//\\\_______\///\\\\\\\\\/____\///\\\\\\\\\/___\/\\\_____________\/\\\_\/\\\_____________\/\\\_\///\\\\\\\\\\\/_______/\\\____ 
        ______\///____\///__________\/////////________\/////////_____\///______________\///__\///______________\///____\///////////________\///_____
"@
        
        Write-Host -ForegroundColor Red $wumms

    } else {}
}

## main function

# do we need more power?
use-power

# create random list
$SampleObject = Import-Excel -Path $SourceFile
$randomlist = for ($i=1; $i -le $SampleSize; $i++) { 
	$SampleObject[(Get-Random -Maximum ($SampleObject.count))]; 
}

# write list
$randomlist | Export-Excel $OutFile -WorksheetName Samplelist -AutoSize -AutoFilter -TableName "Samples"
