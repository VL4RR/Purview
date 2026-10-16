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

$DLPPol = Get-DlpCompliancePolicy -Identity "ComplexTest"

foreach ($DLPPoll in $DLPPol.Name) {
    $DLPRules = Get-DlpComplianceRule -Policy $DLPPoll
    foreach ($DLPRule in $DLPRules.Name) {
        $DLPRuleInfo = Get-DlpComplianceRule -Identity $DLPRule
        Write-Host "$($DLPRuleInfo.DisplayName)" -foregroundcolor Green                
        $DLPAdvRule = $DLPRuleInfo.AdvancedRule | ConvertFrom-Json
        $DLPAdvRuleVal = $DlPAdvRule.Condition.SubConditions
        foreach ($condition in $DLPAdvRuleVal) {
            if ($condition.ConditionName -eq "ContentContainsSensitiveInformation") {
                Write-host "$($allConditionsCondition[$condition.ConditionName])"
                foreach ($sensitiveType in $condition.Value.Groups.Sensitivetypes) {
                    Write-Host "$($sensitiveType.Name) --- $($sensitiveType.ConfidenceLevel) --- $($sensitiveType.MinCount) to $($sensitiveType.MaxCount)"
                }
            }
            else {
                Write-Host "$($allConditionsCondition[$condition.ConditionName])`n$($condition.Value)`n"
            }
        }
    }
}

