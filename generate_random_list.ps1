<#
.SYNOPSIS
	Will generate a random list from a given List.
	
.DESCRIPTION
	In case you need to generate a sample for your internal control system, you could use this script.
	The script then will generate a new Excelfile with the samples for you. Completly random.

.Parameter SourceFile 
	Defines the sourcefile from which the sample is generated.
.Parameter OutFile 
	Defines the outputfile for the sample.
	
.INPUTS
	Needs to be written.

.OUTPUTS
	Needs to be written.

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
## parametersection
param(
	[Parameter(Mandatory=$true,ParameterSetName="SourceFile",HelpMessage="Sourceinputfile for sample creation.")]
	[Alias('source')]
	[string]$SourceFile,
	[Parameter(Mandatory=$false,ParameterSetName="OutFile",HelpMessage="File where the output will be written.")]
	[Alias('output')]
	[string]$OutFile="iks-sample.xlsx"
)

$excelfile = Import-Excel -Path Tickets.xlsx ;
$amount_ticket = 40;

## install & import modules or namespaces
Set-ExecutionPolicy Bypass -Scope Process
Install-Module -Name ImportExcel -scope CurrentUser
Import-Module -Name ImportExcel

## define runtime
$ErrorActionPreference = [ActionPreference]::Stop
Set-StrictMode -Version 'Latest'

## main function

$randomlist = for ($i=1; $i -le $amount_ticket; $i++) { 
	$excelfile[(Get-Random -Maximum ($excelfile.count))]; 
}

$randomlist | Export-Excel

## cleanup (if necessary)

## digital signature
