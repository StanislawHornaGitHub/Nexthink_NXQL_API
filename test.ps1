function invoke-main {
& $testone
    
    Write-Host $script:a
    Write-Host $script:b

}

$testone = {
    $script:a = 1
    $script:b = 2
    Write-Host "Done"
}

invoke-main