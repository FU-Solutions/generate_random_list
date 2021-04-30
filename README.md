# generate_random_list
## Description
In case you need to generate a sample for your internal control system or audit, you could use this script.
The script then will generate a new Excelfile with the samples for you or the auditor. Completly random.

### Parameter SourceFile 
Defines the sourcefile from which the sample is generated.
No default set - please define an input file.
### Parameter OutFile 
Defines the outputfile for the sample.
Default is iks-sample.xlsx.
### Parameter SampleSize
Defines the amount of entries in your sample-list.
Default are 40 samples.
### Parameter poWer
This switch allows you to use more power for your request! Just use switch -w and lay back.
	
### Input
You have to define an excel-list with parameter SourceFile.
In general it will get 40 entries from you List - you can change that

### Outputs
You will get an excel-list with the same data as the inputfile but only the amount of columns you chose.
