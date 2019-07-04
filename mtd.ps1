$targetDir = (Convert-Path .);
$option = $Args[0];
$mdFileList = '';
$itemList = '';
$docxName = 'output.docx';

Write-Host('Now processing...');

function showUsage() {
    Write-Host("usage: mtd [-h | --help] [-r | --recurse]`r`n");
}

If($option -eq $null) {
    $mdFileList = Get-ChildItem $targetDir -Filter *.md;

    If($mdFileList -ne $null) {
        $cmd = 'pandoc ' + $mdFileList + ' -t docx -o output.docx';
        Invoke-Expression $cmd;
        Write-Host('Done');
        Write-Host('Output file is ' + $targetDir + '\' + $docxName);
    }Else {
        Write-Error('ERROR: No .md files in this directory');
    }

}ElseIf($option -eq '-h' -Or $option -eq '--help') {
    showUsage;
    Write-Host('-h | --help'.PadRight(20) + 'show help');
    Write-Host('-r | --recursive'.PadRight(20) + 'convert all .md files under ');
    Write-Host(''.PadRight(20) + 'current directory');

}ElseIf($option -eq '-r' -Or $option -eq '--recurse') {
    $mdFileList = Get-ChildItem $targetDir -include *.md -Recurse | ForEach-Object {$_.FullName};
    $childDirectoryList = @();
    $docxList = '';

    ForEach($item in $mdFileList){
        $mdDirectory = Split-Path $item -parent;
        If(($mdDirectory -ne $targetDir) -And !($childDirectoryList.Contains($mdDirectory))){
            $childDirectoryList += $mdDirectory;
        }
    }

    ForEach($dir in $childDirectoryList){
        cd $dir;
        $currentDir = (Convert-Path .);
        $mdFileList = Get-ChildItem $currentDir -Filter *.md;
        $cmd = 'pandoc ' + $mdFileList + ' -t docx -o ' + $docxName;
        Invoke-Expression $cmd;
        $docxList += $dir + '\output.docx ';
    }

    cd $targetDir;

    If($docxList -ne '') {
        $cmd = 'pandoc ' + $docxList + '-o output.docx';
        Invoke-Expression $cmd;
        $cmd = 'rm ' + $docxList;
        Invoke-Expression $cmd;
        Write-Host('Done');
        Write-Host('Output file is ' + $targetDir + '\' + $docxName);
    }Else {
        Write-Error('ERROR: No .md files under this directory')
    }

}Else {
    Write-Host('unknown option: ' + $option);
    showUsage;
}
