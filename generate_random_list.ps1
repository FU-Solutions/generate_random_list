<#
.SYNOPSIS
	Will generate a random list from a given List.
	
.DESCRIPTION
	
.PARAMETER SourceFile
	Sourceinputfile for sample creation.
 
.PARAMETER OutFile
	File where the output will be written.
	
.NOTES
	File Name     	: generate_random_list.ps1
	Author         	: Ulli Weichert
	Contact	   		: ulli@weichert.it - Weichert.IT
	Version			: 0.1
	
	Introduction:
		In case you need to generate a sample for your internal control system, you could use this script.
        The script then will generate a new Excelfile with the samples for you. Completly random.
        
	Script details:
		Needs to be written.
#>

## install & import modules or namespaces
Set-ExecutionPolicy Bypass -Scope Process
Install-Module -Name ImportExcel -scope CurrentUser
Import-Module -Name ImportExcel

## define runtime
$ErrorActionPreference = [ActionPreference]::Stop
Set-StrictMode -Version 'Latest'

## parameter and variables
$excelfile = Import-Excel -Path Tickets.xlsx ;
$amount_ticket = 40;

param (
	[Parameter(Mandatory=$false,ParameterSetName="Sourcefile",HelpMessage="Sourceinputfile for sample creation.")]
	[Alias('source')]
	[string[]]$SourceFile,
	[Parameter(Mandatory=$false,ParameterSetName="Outputfile",HelpMessage="File where the output will be written.")]
	[Alias('output')]
	[string[]]$OutFile
)

## functions

## main function

$randomlist = for ($i=1; $i -le $amount_ticket; $i++) { $excelfile[(Get-Random -Maximum ($excelfile.count))]; }
$randomlist | Export-Excel

## cleanup (if necessary)

## digital signature