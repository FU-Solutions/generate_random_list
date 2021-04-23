<#
.SYNOPSIS
	Will generate a random sample list - for audits or so - from a given list.
	
.DESCRIPTION
	In case you need to generate a sample for your internal control system or audit, you could use this script.
	The script then will generate a new Excelfile with the samples for you or the auditor. Completly random.

.Parameter SourceFile 
	Defines the sourcefile from which the sample is generated.
.Parameter OutFile 
	Defines the outputfile for the sample.
.Parameter SampleSize
	Defines the amount of entries in your sample-list.
	
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
	Version			: 0.1
	
	Introduction:
		In case you need to generate a sample for your internal control system, you could use this script.
		The script then will generate a new Excelfile with the samples for you. Completly random.
        
	Script details:
		Needs to be written.

.LINK
	https://github.com/w3ich3rt/generate_random_list
#>

## Parameters and variables

param(
	[Parameter(Mandatory=$false,ValueFromPipeline=$true,ParameterSetName="SourceFile",HelpMessage="Sourceinputfile for sample creation.")]
	[Alias('source')]
	[string]$SourceFile,
	[Parameter(Mandatory=$false,ParameterSetName="OutFile",HelpMessage="File where the output will be written.")]
	[Alias('output')]
	[string]$OutFile="iks-sample.xlsx",
	[Parameter(Mandatory=$false,ParameterSetName="Samples",HelpMessage="Amount of sample in you generated list")]
	[Alias('samples')]
	[Int]$SampleSize="40"
)

$SampleObject = Import-Excel -Path $SourceFile

## install & import modules or namespaces
Install-Module -Name ImportExcel -scope CurrentUser
Import-Module -Name ImportExcel

## main function

$randomlist = for ($i=1; $i -le $SampleSize; $i++) { 
	$SampleObject[(Get-Random -Maximum ($SampleObject.count))]; 
}

$randomlist | Export-Excel $OutFile -WorksheetName Samplelist -AutoSize -AutoFilter
