param([string]$StdOutPath)

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");

[string]$PSDefaultParameterValues['*:Encoding']        = 'utf8';
[string]$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8';
[System.Windows.Forms.SaveFileDialog]$dialog   		   = [System.Windows.Forms.SaveFileDialog]::new();
[string]$dialog.Filter                                 = "HTML files (*.html)|*.html|EXEL files (*.xlsx)|*.xlsx";
[bool]$dialog.RestoreDirectory                         = $true;
[object]$result                                        = $dialog.ShowDialog();

if ($result -eq "OK")
{
   echo $dialog.FileName > $StdOutPath;
} else {
   echo 1                > $StdOutPath;
}
