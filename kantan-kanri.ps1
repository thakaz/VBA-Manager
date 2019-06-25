param([string]$fileName, [string]$cmd)  
#$fileName = "hoge.xlsm"
#$cmd = "export"


$targetFolder = $PSScriptRoot 

$srcFolderPath = Join-Path $targetFolder "src"

$binFolderPath = Join-Path $targetFolder "bin"
$tmpFolderPath = Join-Path $targetFolder "template"

$binFilePath = Join-path $binFolderPath $fileName
$tmpFilepath = Join-path $tmpFolderPath $fileName 

#���W���[���̎��
$moduleTypeTable = @{
  "1" = ".bas" 
  "2"   = ".cls"
  "3" = ".frm"
  "100" = ".bas"
}
  
function CreateTemplate {

  New-Item $tmpFolderPath -ItemType Directory -Force

  Copy-Item $binfilePath  $tmpFilepath -Force

  $excel = New-Object -ComObject Excel.Application  
  $book = $excel.Workbooks.Open($tmpFilepath) 
  
  #�܂��V�[�g���W���[�����폜���Ȃ��ƃ��[�U�[��`�^�]�X�ł����炭�{����
  $book.VBProject.VBComponents | % { 
    If ($_.Type -eq "100") { 
      $_.CodeModule.DeleteLines(1, $_.CodeModule.CountOfLines)
    }
  }  
  #���ׂẴ��W���[�����폜����ƂȂ����Q�Ɛݒ肪�����邽�߁A�_�~�[���c��
  $book.VBProject.VBComponents | % { 
    If ($_.Type -in @("1", "2", "3") -and $_.Name -ne "Dummy") { 
      $book.VBProject.VBComponents.Remove($_)  
    }
  }  

  $book.Save()
  $book.Close(0)
  $excel.Quit()  
  $excel = $null
  [System.GC]::Collect()
}

function ExportModule {

  Remove-Item Join-Path $srcFolderPath   -Force -Recurse
  New-Item $srcFolderPath -ItemType Directory -Force 
  New-Item (JOIN-Path $srcFolderPath "sht") -ItemType Directory -Force 
  #New-Item (JOIN-Path $srcFolderPath "std") -ItemType Directory -Force 
  
  $excel = New-Object -ComObject Excel.Application  
    
  $excel.Workbooks.Open($binFilePath) | % {  
    $_.VBProject.VBComponents | % { 

      If ($_.Type -eq "100") {
        $exportFileName = JOIN-Path $srcFolderPath "sht" | Join-Path -Childpath ($_.Name + $moduleTypeTable[[string]$_.Type])  
      }
      else {
        $exportFileName = JOIN-Path $srcFolderPath "" | Join-Path -Childpath ($_.Name + $moduleTypeTable[[string]$_.Type])      
      }
      $_.Export($exportFileName)  
    }  

    $_.Close(0)
  }  
    
  $excel.Quit()  
  $excel = $null
  [System.GC]::Collect()

}

function ImportModule {

  Remove-Item $binFolderPath  -Force -Recurse
  New-Item $binFolderPath -ItemType Directory -Force 

  Copy-Item $tmpFilepath  $binFilePath -Force
  
  $excel = New-Object -ComObject Excel.Application  
  $book = $excel.Workbooks.Open($binFilePath)

  Get-childItem -Path $srcFolderPath -File | % {
    #export����frm�`���̃t�@�C�����o�͂����̂ł����import�ΏۊO�Ƃ���

    If ($_.Extension -ne ".frx") {
      $book.VBProject.VBComponents.Import($_.FullName)
    }
  }
  $shtFolder = JOIN-Path $srcFolderPath "sht"
  Get-childItem -Path $shtFolder -File | % {
    $sheetName = [System.IO.Path]::GetFileNameWithoutExtension($_) 
    $tmpVBE = $book.VBProject.VBComponents($sheetName).CodeModule
    
    $tmpVBE.DeleteLines(1, $tmpVBE.CountOfLines)
    $tmpVBE.AddFromFile($_.FullName)
    #�擪4�s�͕s�v
    $tmpVBE.DeleteLines(1, 4)

  }

  
  $book.Save()
  $book.Close(0) 
  $excel.Quit()  
  $excel = $null
  [System.GC]::Collect()

  
}


#���C���J�n
switch ($cmd) {
  
  "export" {
    ExportModule
    CreateTemplate    
  }
  "import" {
    ImportModule
  }
  "clear" {
    CreateTemplate
  }
  "exportonly" {
    ExportModule
  }
  default {
    Add-Type -Assembly System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("����������", "������", "OK", "Warning", "button1")
  }
}

echo "END"

