function WindowsDialogBox {
    param (
        [Parameter(Mandatory=$true)]
        [string]$MessageBody,

        [Parameter(Mandatory=$true)]
        [string]$MessageTitle
    )

    Add-Type -AssemblyName PresentationCore,PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::YesNoCancel
    $MessageIcon = [System.Windows.MessageBoxImage]::Question
    $DialogResult = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
    Write-Host $DialogResult
}
