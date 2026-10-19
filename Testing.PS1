$allConditionsCondition = @{
    "ContentContainsSensitiveInformation" = "Content Contains:"
    "AccessScope"                         = "Content is shared:"
    "SenderIPRanges"                      = "Sender IP address is:"
    "RecipientDomainIs"                   = "Recipient domain is:"
    "ContentExtensionMatchesWords"        = "Attachment extension is:"
    "DocumentIsPasswordProtected"         = "Attachment is password protected"
    "DocumentIsUnsupported"               = "Any email attachment's content could not be scanned"
    "ProcessingLimitExceeded"             = "Any email attachment's content didn't complete scanning"
    "HeaderMatchesPatterns"               = "Header matches patterns:"
    "SubjectMatchesPatterns"              = "Subject matches patterns:"
    "AnyOfRecipientAddressContainsWords"  = "Recipient address contains words:"
    "FromAddressMatchesPatterns"          = "Sender address matches pattern"
    "SenderDomainIs"                      = "Sender domain is:"
    "SentToMemberOf"                      = "Recipient is a member of:"
    "From"                                = "Sender is:"
    "DocumentNameMatchesWords"            = "Attachment name matches words or phrases:"
    "ContentPropertyContainsWords"        = "Attachment property is:"
    "DocumentSizeOver"                    = "Attachment size equals or is greater than:"
    "SentTo"                              = "Recipient is:"

    "DocumentCreatedBy"                   = "Document created by:"
    "DocumentCreatedOnOrAfter"            = "Document creation date is on or after"
    "DocumentCreatedOnOrBefore"           = "Document creation date is on or before"
    "DocumentLastModifiedOnOrAfter"       = "Document modification date is on or after"
    "DocumentLastModifiedOnOrBefore"      = "Document modification date is on or before"
}

$Labels = Get-Label
$script:Labels = $Labels

function SensitiveTypes {
    param($Condition)

    $sensitiveTypeTC = $Condition.Value.Groups.Sensitivetypes | Where-Object { $_.Classifiertype -eq "MLModel" -and $_.Name }
    $sensitiveTypeSIT = $Condition.Value.Groups.Sensitivetypes | Where-Object { $_.Classifiertype -ne "MLModel" -and $_.Name }
    $sensitiveTypeLabel = $Condition.Value.Groups.Labels | Where-Object { $_.Name }
    Write-Host "$($Condition.ConditionName)`n"
    if ($sensitiveTypeTC) {
        Write-Host "Trainable Classifiers:"
        foreach ($sensitiveType in $sensitiveTypeTC) { Write-Host "$($sensitiveType.Name)" }
        Write-Host ""
    }
    if ($sensitiveTypeSIT) {
        Write-Host "Sensitive Information Types:"
        foreach ($sensitiveType in $sensitiveTypeSIT) {
            Write-Host "$($sensitiveType.Name) --- $($sensitiveType.ConfidenceLevel) --- $($sensitiveType.MinCount) to $($sensitiveType.MaxCount)"
        }
        Write-Host ""
    }
    if ($sensitiveTypeLabel) {
        Write-Host "Labels:"
        foreach ($sensitiveType in $sensitiveTypeLabel) {
            $label = $script:Labels | Where-Object { $_.Guid -eq $sensitiveType.Id } | Select-Object -First 1
            if ($label) { Write-Host $label.DisplayName }
            else { Write-Host $sensitiveType.Name }
        }
        Write-Host ""
    }
}

function HeaderPatterns {
    param($Condition)
    Write-Host "Header Matches Patterns:"
    Write-host "Header name: $($Condition.Value.PSObject.Properties.Name)`nHeader value:`n$($Condition.Value.PSObject.Properties.Value -join "`n")`n"
}

$DLPPol = Get-DlpCompliancePolicy -Identity "ComplexTest"

foreach ($DLPPoll in $DLPPol.Name) {
    $DLPRules = Get-DlpComplianceRule -Policy $DLPPoll
    foreach ($DLPRule in $DLPRules.Name) {
        $DLPRuleInfo = Get-DlpComplianceRule -Identity $DLPRule
        Write-Host "$($DLPRuleInfo.DisplayName)`n" -ForegroundColor Green
        $DLPAdvRule = $DLPRuleInfo.AdvancedRule | ConvertFrom-Json
        $DLPAdvRuleVal = $DLPAdvRule.Condition.SubConditions
        if ($DLPAdvRuleVal.Operator.Count -ge 1) { 
            Write-Host "Group 1:`n"  
            $groupNumber = 1 
        }
        foreach ($condition in $DLPAdvRuleVal) {
            if ($condition.ConditionName -eq "ContentContainsSensitiveInformation") { SensitiveTypes $condition }
            elseif ($condition.ConditionName -eq "HeaderMatchesPatterns") { HeaderPatterns $condition }
            elseif ($condition.Operator -eq "And") {
                $groupNumber ++
                Write-Host "Group $groupNumber`n"
                foreach ($GroupCondition in $condition.SubConditions) {
                    if ($GroupCondition.ConditionName -eq "ContentContainsSensitiveInformation") { SensitiveTypes $GroupCondition }
                    elseif ($GroupCondition.ConditionName -eq "HeaderMatchesPatterns") { HeaderPatterns $GroupCondition }
                    else { Write-Host "$($GroupCondition.ConditionName)`n$($GroupCondition.Value)`n" }
                }
            }
            else { Write-Host "$($condition.ConditionName)`n$($condition.Value)`n" }
        }
    }
}