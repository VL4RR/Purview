$DLPPol = Get-DlpCompliancePolicy 

foreach ($DLPPoll in $DLPPol.Name) {
    $DLPRules = Get-DlpComplianceRule -Policy $DLPPoll
    Write-Host "These are the rules for $DLPPoll" -ForegroundColor Blue
    foreach ($DLPRule in $DLPRules.Name) {
        $DLPRuleInfo = Get-DlpComplianceRule -Identity $DLPRule
        $DLPRuleInfo.Name
        $DLPRuleInfo
        Write-Host ""

    }
    Write-Host ""
}