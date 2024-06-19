<#
.SYNOPSIS
This command binds several pdf documents together
 
.DESCRIPTION
This command binds several pdf documents together
I included this command in the module, because otherwise
I didn't find anything for PowerShell. It will surely help one or the other.
 
.PARAMETER fileNames
A List of PDF Files zu combine
 
.PARAMETER OutputPdfDocument
The path to the output file
 
.EXAMPLE
New-CombineMultiplePDFs -fileNames @('c:\temp\file1.pdf','c:\temp\file2.pdf') -OutputPdfDocument 'c:\temp\combined.pdf'
 

#>

function New-CombineMultiplePDFs
{
  param(
    [string[]] $fileNames, 
    [System.IO.FileInfo] $OutputPdfDocument
  )
  
  if (test-path "$OutputPdfDocument") { Remove-Item "$OutputPdfDocument"  }
  
  $fileStream = New-Object System.IO.FileStream($OutputPdfDocument, [System.IO.FileMode]::OpenOrCreate)
  $document = New-Object iTextSharp.text.Document
  $pdfCopy = New-Object iTextSharp.text.pdf.PdfCopy($document, $fileStream)
    
   
  $document.Open()
  
  foreach ($fileName in $fileNames)
  {
    [System.IO.FileInfo] $fi = $fileName
    $reader = New-Object iTextSharp.text.pdf.PdfReader -argumentlist $fi.fullname
    $reader.ConsolidateNamedDestinations()
    #$pdfCopy.AddDocument($reader);
    
                for ($i = 1; $i -le $reader.NumberOfPages; $i++)
                {
                    $page =  $pdfCopy.GetImportedPage($reader, $i);
                    $pdfCopy.AddPage($page);
                }



    $reader.Close()
  }

  $pdfCopy.Close();
  $document.Close();
  $fileStream.Close(); 
}
